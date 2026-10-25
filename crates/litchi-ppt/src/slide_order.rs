//! Unified source-checked editing for opened legacy PPT presentations.
//!
//! `[MS-PPT]` 2.4.14.3 defines presentation order by the sequence of
//! `SlidePersistAtom` records in `SlideListWithTextContainer`. This owner moves
//! each complete slide group, including the outline-text records following its
//! persist atom, and publishes the changed live `DocumentContainer` through a
//! new append-only PPT user edit. The same root composes checked shape text
//! and complete PowerPoint client-anchor geometry formatting. Slide payload
//! records and unrelated CFB streams remain exact unless selected explicitly.
//!
//! The editing owner (its transaction, commit and patch) is crate-private
//! since record 0357: [`crate::Package::edit`] is the one public edit route.
//! This module keeps public the vocabulary those verbs take, the typed
//! [`Error`](crate::slide_order::Error)/[`Refusal`](crate::slide_order::Refusal)
//! they return, and [`Snapshot`](crate::slide_order::Snapshot), an immutable
//! whole-artifact read view of the slide list and its slide-level state.
//! Its durable patches, patch joins and three-way planning are reached
//! through `package::edit::Patch::{to_durable, join}` and
//! `Package::{apply_durable, plan_three_way}` (record 0357a).
#![allow(
    dead_code,
    reason = "crate-private native owner of the primary edit (record 0357); its durable, join and three-way surfaces are public through package::edit (record 0357a), the rest stays exercised by the unit tests"
)]

use std::collections::{BTreeMap, BTreeSet};
use std::fmt;
use std::io::Cursor;
use std::sync::Arc;

pub use litchi_core::Position;

use crate::document_structure;
use crate::package::{Error as PackageError, OpenOptions, SourceLimits};
use crate::raw::Reader;

/// Maximum automatic slide-advance delay permitted by `[MS-PPT]` 2.6.6.
pub const MAX_SLIDE_ADVANCE_MS: u32 = 86_399_000;

/// Maximum shape-text replacements accepted by one atomic batch.
///
/// The bound is checked from the slice length before target resolution,
/// package parsing, or replacement encoding.
pub const MAX_SHAPE_TEXT_REPLACEMENTS: usize = 256;

/// Maximum number of chunks accepted by one aggregate shape-text operation.
///
/// The bound includes empty chunks. It is checked before the source
/// presentation or any selected target is resolved.
pub const MAX_SHAPE_TEXT_CHUNKS: usize = 4096;

/// Maximum number of globally selected shapes accepted by one aggregate
/// shape-text operation.
///
/// Raised from 4096 with the primary edit bound (record 0264, owner approved
/// 2026-10-14): verification no longer re-parses the slide, so one larger
/// publication is linear where several smaller ones each rewrote the slide.
pub const MAX_SHAPE_TEXT_CHUNK_SELECTIONS: usize = 16384;

/// Fixed-width manual and automatic advance state from one
/// `SlideShowSlideInfoAtom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SlideAdvance {
    manual: bool,
    automatic: bool,
    delay_ms: u32,
}

impl SlideAdvance {
    /// Creates a validated advance state.
    ///
    /// The stored delay remains explicit even when automatic advance is off,
    /// allowing exact inverse restoration of producer-authored ignored bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when `delay_ms` exceeds the binary-format maximum.
    pub fn new(manual: bool, automatic: bool, delay_ms: u32) -> Result<Self> {
        if delay_ms > MAX_SLIDE_ADVANCE_MS {
            return Err(PackageError::InvalidFormat(format!(
                "PPT slide advance delay {delay_ms} exceeds {MAX_SLIDE_ADVANCE_MS}"
            ))
            .into());
        }
        Ok(Self {
            manual,
            automatic,
            delay_ms,
        })
    }

    /// Creates click-only advance state with a zero stored delay.
    #[must_use]
    pub const fn on_click() -> Self {
        Self {
            manual: true,
            automatic: false,
            delay_ms: 0,
        }
    }

    /// Creates automatic-only advance state.
    ///
    /// # Errors
    ///
    /// Returns an error when `delay_ms` exceeds the format maximum.
    pub fn automatic(delay_ms: u32) -> Result<Self> {
        Self::new(false, true, delay_ms)
    }

    /// Creates click-or-automatic advance state.
    ///
    /// # Errors
    ///
    /// Returns an error when `delay_ms` exceeds the format maximum.
    pub fn both(delay_ms: u32) -> Result<Self> {
        Self::new(true, true, delay_ms)
    }

    /// Whether the user may manually advance the slide.
    #[must_use]
    pub const fn manual(self) -> bool {
        self.manual
    }

    /// Whether the slide advances automatically.
    #[must_use]
    pub const fn automatic_enabled(self) -> bool {
        self.automatic
    }

    /// Stored automatic delay in milliseconds.
    #[must_use]
    pub const fn delay_ms(self) -> u32 {
        self.delay_ms
    }
}

/// Fixed-width visual transition state from one `SlideShowSlideInfoAtom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SlideTransitionVisual {
    transition_type: crate::TransitionType,
    direction: crate::TransitionDirection,
    speed: crate::TransitionSpeed,
    wire: [u8; 3],
}

impl SlideTransitionVisual {
    /// Creates a transition state that has an exact legacy PPT wire mapping.
    ///
    /// # Errors
    ///
    /// Returns an error when the type/direction combination is not represented
    /// exactly by the established `SlideShowSlideInfoAtom` mappings.
    pub fn new(
        transition_type: crate::TransitionType,
        direction: crate::TransitionDirection,
        speed: crate::TransitionSpeed,
    ) -> Result<Self> {
        let Some(wire) = crate::transition::encode_visual(transition_type, direction, speed) else {
            return Err(PackageError::InvalidFormat(
                "PPT visual transition type/direction combination is not exact".into(),
            )
            .into());
        };
        Ok(Self {
            transition_type,
            direction,
            speed,
            wire,
        })
    }

    /// Transition effect kind.
    #[must_use]
    pub const fn transition_type(self) -> crate::TransitionType {
        self.transition_type
    }

    /// Effect direction, when the selected kind has one.
    #[must_use]
    pub const fn direction(self) -> crate::TransitionDirection {
        self.direction
    }

    /// Effect playback speed.
    #[must_use]
    pub const fn speed(self) -> crate::TransitionSpeed {
        self.speed
    }
}

/// A native owner edge that must be closed before a slide can be transferred.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum TransferDependency {
    /// The slide points at a notes persist object.
    SpeakerNotes,
    /// Slide comments also require the document author catalog.
    CommentAuthorCatalog,
    /// Interactive actions require the document hyperlink relationship table.
    HyperlinkAction,
    /// `OfficeArt` shape identifiers and drawing-group state are
    /// presentation-global.
    DrawingGroup,
    /// A connector rule refers to a shape outside the transferred slide's
    /// closed drawing identity set.
    ConnectorRule,
    /// Another `OfficeArt` shape-reference family (arc/callout/deleted-shape
    /// rule, drawing selection, or linked-shape property) is present.
    OtherOfficeArtShapeReference,
    /// BLIP-bearing `OfficeArt` properties require a closed picture-store copy
    /// and index remap that this transfer owner does not perform.
    PictureStore,
    /// Animation and build records address shape identities through graphs
    /// whose complete fixed-width reference closure is not modeled here.
    AnimationBuildGraph,
    /// An `ExObjRefAtom` can address media, a chart, or another OLE object in
    /// the document external-object relationship table.
    ExternalObjectRelationship,
    /// An interaction or external-object owner contains executable, active
    /// OLE, or unknown relationship semantics that are never transferred.
    ActiveOrUnknownExternalObject,
}

impl fmt::Display for TransferDependency {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::SpeakerNotes => "speaker-notes",
            Self::CommentAuthorCatalog => "comment-author-catalog",
            Self::HyperlinkAction => "hyperlink-action",
            Self::DrawingGroup => "drawing-group",
            Self::ConnectorRule => "open connector rule",
            Self::OtherOfficeArtShapeReference => "other OfficeArt shape reference",
            Self::PictureStore => "picture-store/BLIP",
            Self::AnimationBuildGraph => "animation/build graph",
            Self::ExternalObjectRelationship => "external media/chart/OLE relationship",
            Self::ActiveOrUnknownExternalObject => "active or unknown external object",
        })
    }
}

/// A reason why slide order cannot safely be changed for this source.
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Refusal {
    /// The CFB is encrypted or contains storage metadata that the stream-only
    /// incremental publisher cannot preserve faithfully. Signed sources have
    /// the more precise [`Refusal::Signed`] policy below.
    UnsupportedSource,
    /// The CFB carries a digital-signature component.  Read-only planning and
    /// exact no-op transactions remain valid, but any changed publication
    /// would invalidate the signature and is refused before a lower-level
    /// writer can drop or rewrite it.
    Signed,
    /// The read path repaired a cyclic UserEdit chain.  The native append-only
    /// editor still requires a strict chain, so this source is read-only.
    RepairedUserEditChain,
    /// A source or destination position is outside the current slide list.
    SlideNotFound { position: Position },
    /// Reviewer diff trees associate their slide entries by list position.
    ReviewHistoryDependency,
    /// A slide depends on notes, drawings, media, or another owner whose
    /// identifiers cannot yet be rewritten as a closed transfer unit.
    UnsupportedSlideDependency { dependency: TransferDependency },
    /// The receiving presentation does not contain the referenced master.
    MissingMasterDependency { master_id: u32 },
    /// The receiver reuses the master ID for different persisted content.
    MismatchedMasterDependency { master_id: u32 },
    /// A transfer plan was prepared for a different exact receiving artifact.
    TransferTargetMismatch,
    /// The selected slide does not contain exactly one valid fixed-width
    /// `SlideShowSlideInfoAtom` to mutate without record insertion/removal.
    UnsupportedSlideAdvance { position: Position },
    /// The selected slide does not contain exactly one valid fixed-width
    /// `SlideShowSlideInfoAtom` whose `fHidden` flag can be changed in place.
    UnsupportedSlideVisibility { position: Position },
    /// The selected slide does not contain exactly one canonical fixed-width
    /// visual transition owner.
    UnsupportedSlideTransitionVisual { position: Position },
    /// A shape edit selected a slide inserted but not yet published.
    UncommittedSlideDependency,
    /// One semantic shape target appeared more than once in one batch.
    DuplicateShapeTextTarget { target: crate::text_edit::Target },
    /// An element edit (record 0323) would need records whose layout it does
    /// not model losslessly.
    UnsupportedElement {
        element: &'static str,
        reason: &'static str,
    },
    /// An element edit selected an element that does not exist.
    ElementNotFound { element: &'static str },
    /// An element edit was given a value its records cannot hold.
    InvalidElementValue {
        element: &'static str,
        reason: &'static str,
    },
    /// A patch operation (a join, or a durable conversion) is not supported
    /// for the content of this patch (record 0357a).
    UnsupportedPatch { reason: &'static str },
    /// Two joined patches write the same semantic target (record 0357b),
    /// named as a three-way plan names it (for example
    /// `document/master-list`), so their order would decide the result.
    JoinConflict { target: String },
}

impl fmt::Display for Refusal {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::UnsupportedSource => formatter.write_str(
                "PPT source cannot be republished through the safe slide-order transaction",
            ),
            Self::Signed => formatter.write_str(
                "PPT signed source cannot be changed through the safe slide-order transaction",
            ),
            Self::RepairedUserEditChain => formatter
                .write_str("PPT source contains a repaired UserEdit chain and is read-only"),
            Self::SlideNotFound { position } => {
                write!(
                    formatter,
                    "PPT slide position {} was not found",
                    position.get()
                )
            },
            Self::ReviewHistoryDependency => formatter
                .write_str("PPT slide order is referenced by unmodeled reviewer diff history"),
            Self::UnsupportedSlideDependency { dependency } => write!(
                formatter,
                "PPT slide transfer has an unsupported {dependency} dependency"
            ),
            Self::MissingMasterDependency { master_id } => write!(
                formatter,
                "PPT transfer target has no matching master {master_id:#010x}"
            ),
            Self::MismatchedMasterDependency { master_id } => write!(
                formatter,
                "PPT transfer target master {master_id:#010x} has different persisted content"
            ),
            Self::TransferTargetMismatch => {
                formatter.write_str("PPT slide transfer plan belongs to a different target")
            },
            Self::UnsupportedSlideAdvance { position } => write!(
                formatter,
                "PPT slide position {} has no unique fixed-width slide-show information atom",
                position.get()
            ),
            Self::UnsupportedSlideVisibility { position } => write!(
                formatter,
                "PPT slide position {} has no unique fixed-width slide-show visibility owner",
                position.get()
            ),
            Self::UnsupportedSlideTransitionVisual { position } => write!(
                formatter,
                "PPT slide position {} has no unique canonical visual transition atom",
                position.get()
            ),
            Self::UncommittedSlideDependency => formatter
                .write_str("PPT shape text cannot target an inserted slide before publication"),
            Self::DuplicateShapeTextTarget { target } => write!(
                formatter,
                "PPT shape-text batch repeats slide {}/shape {}",
                target.slide().get(),
                target.shape().get()
            ),
            Self::UnsupportedElement { element, reason } => {
                write!(formatter, "PPT {element} edit is not supported: {reason}")
            },
            Self::ElementNotFound { element } => {
                write!(formatter, "PPT {element} was not found")
            },
            Self::InvalidElementValue { element, reason } => {
                write!(formatter, "PPT {element} value is invalid: {reason}")
            },
            Self::UnsupportedPatch { reason } => {
                write!(formatter, "PPT patch operation is not supported: {reason}")
            },
            Self::JoinConflict { target } => {
                write!(formatter, "PPT joined patches both write {target}")
            },
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
            Self::Signed => RefusalKind::Protected,
            Self::SlideNotFound { .. } | Self::ElementNotFound { .. } => RefusalKind::NotFound,
            Self::DuplicateShapeTextTarget { .. } | Self::InvalidElementValue { .. } => {
                RefusalKind::InvalidValue
            },
            Self::MissingMasterDependency { .. }
            | Self::MismatchedMasterDependency { .. }
            | Self::TransferTargetMismatch
            | Self::UnsupportedSlideAdvance { .. }
            | Self::UnsupportedSlideVisibility { .. }
            | Self::UnsupportedSlideTransitionVisual { .. } => RefusalKind::UnsupportedTarget,
            Self::UncommittedSlideDependency
            | Self::UnsupportedPatch { .. }
            | Self::JoinConflict { .. } => RefusalKind::UnsupportedOperation,
            Self::UnsupportedSource
            | Self::RepairedUserEditChain
            | Self::ReviewHistoryDependency
            | Self::UnsupportedSlideDependency { .. }
            | Self::UnsupportedElement { .. } => RefusalKind::UnsupportedContent,
        }
    }
}

/// Error returned by the opened-presentation slide-order owner.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The package could not be opened, validated, or republished.
    Package(PackageError),
    /// The requested operation is not proven lossless for this source.
    Refused(Refusal),
    /// A prepared sub-edit exceeded the common composition bounds.
    Composition(litchi_core::patch::CompositionError),
    /// The checked existing-shape text owner rejected or failed an edit.
    Text(Box<crate::text_edit::Error>),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package(error) => error.fmt(formatter),
            Self::Refused(refusal) => refusal.fmt(formatter),
            Self::Composition(error) => error.fmt(formatter),
            Self::Text(error) => error.fmt(formatter),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Package(error) => Some(error),
            Self::Composition(error) => Some(error),
            Self::Text(error) => Some(error.as_ref()),
            Self::Refused(_) => None,
        }
    }
}

impl From<PackageError> for Error {
    fn from(error: PackageError) -> Self {
        Self::Package(error)
    }
}

impl From<litchi_core::patch::CompositionError> for Error {
    fn from(error: litchi_core::patch::CompositionError) -> Self {
        Self::Composition(error)
    }
}

impl From<crate::text_edit::Error> for Error {
    fn from(error: crate::text_edit::Error) -> Self {
        Self::Text(Box::new(error))
    }
}

/// Result type for source-checked slide-order operations.
pub type Result<T> = std::result::Result<T, Error>;

/// A deterministic phase boundary exposed by the opt-in performance
/// diagnostics seam.
///
/// The phases describe format-owned work at the existing commit expression
/// boundaries. They carry no source content, physical identifiers, byte
/// offsets, clocks, or other ambient runtime state.
#[cfg(feature = "performance-diagnostics")]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DiagnosticPhase {
    /// Commits the staged live-document transaction.
    DocumentCommit,
    /// Opens the append-only package editor over the working source.
    EmbeddedOpen,
    /// Reads the persisted slide payloads through the opened package editor
    /// before any record is staged.
    BeforePayloadCapture,
    /// Reads the live document record and checks source consistency before
    /// replacement.
    LiveDocumentRead,
    /// Finishes the append-only package editor.
    EmbeddedFinish,
    /// Checks that unrelated package streams were preserved.
    UnrelatedStreamValidation,
    /// Reopens the finished package through the public snapshot owner.
    PublicReopen,
    /// Captures persisted slide payloads after public reopen. The expected
    /// payload comparison remains explicit residual commit work.
    AfterPayloadCapture,
    /// Computes the SHA-256 of the intermediate artifact that structural
    /// publication started from.
    ///
    /// Only a transaction that stages both formatting and structural changes
    /// reports this phase: the patch retains its source and target artifacts
    /// but not that intermediate, so its digest is computed at commit. Every
    /// other artifact digest is computed on demand by durable serialization
    /// rather than by commit.
    IntermediateArtifactHash,
    /// The staged live-document commit was an exact structural no-op; other
    /// formatting owners may still have staged changes in the root
    /// transaction.
    StructuralNoOp,
}

/// Result class attached to a completed diagnostic phase.
#[cfg(feature = "performance-diagnostics")]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DiagnosticOutcome {
    /// The phase completed and its enclosing operation continued.
    Success,
    /// The phase returned an ordinary typed operation error.
    Error,
}

/// One ordered, content-free phase event emitted by a profiled operation.
///
/// Every started phase has one matching finished event, including phases that
/// return an error. The observer is called synchronously in operation order;
/// the library neither timestamps events nor consults ambient runtime state.
#[cfg(feature = "performance-diagnostics")]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DiagnosticEvent {
    /// A phase is about to run.
    Started { phase: DiagnosticPhase },
    /// A phase has returned to its caller.
    Finished {
        phase: DiagnosticPhase,
        outcome: DiagnosticOutcome,
    },
}

#[cfg(feature = "performance-diagnostics")]
fn observe_phase<T, Operation>(
    observer: &mut impl FnMut(DiagnosticEvent),
    phase: DiagnosticPhase,
    operation: Operation,
) -> Result<T>
where
    Operation: FnOnce() -> Result<T>,
{
    observer(DiagnosticEvent::Started { phase });
    let result = operation();
    observer(DiagnosticEvent::Finished {
        phase,
        outcome: if result.is_ok() {
            DiagnosticOutcome::Success
        } else {
            DiagnosticOutcome::Error
        },
    });
    result
}

mod artifact {
    //! Exact artifact bytes bound to the lazily memoized SHA-256 of exactly
    //! those bytes.
    //!
    //! The fields are private to this module and [`Artifact::new`] is the only
    //! constructor, so new bytes always start with a fresh memo while a clone
    //! keeps its bytes and memo together. No code outside this module can pair
    //! a memo with different bytes, including through struct-update syntax on
    //! an owning snapshot.

    use std::fmt;
    use std::sync::{Arc, OnceLock};

    /// One immutable artifact and its content address.
    ///
    /// The digest is computed at most once per artifact, from exactly its
    /// bytes, and only when a durable-patch or transfer consumer asks for it.
    /// Commit and in-memory patch application never need it: they authorize
    /// by exact byte equality. Every clone shares both the byte allocation and
    /// the memo, a 48-byte allocation that also holds the 64-byte hex digest
    /// once it is computed.
    ///
    /// Equality compares the bytes only, because the digest is a pure function
    /// of them. `Debug` omits the memo, whose fill state depends only on call
    /// history.
    #[derive(Clone)]
    pub(super) struct Artifact {
        bytes: Arc<[u8]>,
        sha256: Arc<OnceLock<String>>,
    }

    impl Artifact {
        /// Binds new exact bytes to a fresh, empty digest memo.
        pub(super) fn new(bytes: Arc<[u8]>) -> Self {
            Self {
                bytes,
                sha256: Arc::default(),
            }
        }

        /// The exact artifact bytes.
        pub(super) fn bytes(&self) -> &[u8] {
            &self.bytes
        }

        /// The shared immutable byte allocation.
        pub(super) const fn shared(&self) -> &Arc<[u8]> {
            &self.bytes
        }

        /// Whether both values share one byte allocation.
        pub(super) fn same_allocation(&self, other: &Self) -> bool {
            Arc::ptr_eq(&self.bytes, &other.bytes)
        }

        /// Lowercase hexadecimal SHA-256 of the bytes, computed on first use.
        pub(super) fn sha256(&self) -> &str {
            self.sha256
                .get_or_init(|| super::artifact_hash(&self.bytes))
        }

        /// The digest if some consumer has already computed it.
        #[cfg(test)]
        pub(super) fn memoized_sha256(&self) -> Option<&str> {
            self.sha256.get().map(String::as_str)
        }

        /// Whether both values share one digest memo.
        #[cfg(test)]
        pub(super) fn shares_memo_with(&self, other: &Self) -> bool {
            Arc::ptr_eq(&self.sha256, &other.sha256)
        }
    }

    impl PartialEq for Artifact {
        fn eq(&self, other: &Self) -> bool {
            self.bytes == other.bytes
        }
    }

    impl Eq for Artifact {}

    impl fmt::Debug for Artifact {
        fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
            formatter
                .debug_struct("Artifact")
                .field("bytes", &self.bytes)
                .finish_non_exhaustive()
        }
    }
}

use artifact::Artifact;

mod blank;
mod elements;
pub(crate) mod records;

// Snapshots memoize their artifact digest through interior mutability
// (`OnceLock`). Keep every shareable slide-order value `Send + Sync`; this
// fails to compile if a future field breaks that.
const _: () = {
    const fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Snapshot>();
    assert_send_sync::<Patch>();
    assert_send_sync::<Commit>();
    assert_send_sync::<TransferPlan>();
};

/// Immutable exact whole-package read view of the slide list and its
/// slide-level state; edits go through [`crate::Package::edit`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    artifact: Artifact,
    document: document_structure::Snapshot,
    document_persist_id: u32,
    limits: SourceLimits,
    leniency: crate::leniency::Leniency,
    #[cfg(feature = "vba-inspection")]
    vba_project_limits: crate::VbaProjectLimits,
    chain_repaired: bool,
    has_review_history: bool,
    has_signature: bool,
}

impl Snapshot {
    /// Opens an exact PPT artifact under default finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the complete package or its live slide directory
    /// is malformed.
    pub(crate) fn parse(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes(bytes.to_vec())
    }

    /// Captures an owned exact PPT artifact under default finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the complete package cannot be reopened.
    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_options(bytes, OpenOptions::new())
    }

    /// Captures an owned exact PPT artifact under explicit source limits and
    /// a selected leniency policy.
    ///
    /// Strict remains the default.  The selected policy is retained by the
    /// immutable snapshot and is reused by every native publication reopen,
    /// so a canonical package edit cannot silently fall back to strict while
    /// processing a caller-selected source.
    pub fn from_bytes_with_options(bytes: Vec<u8>, options: OpenOptions<'_>) -> Result<Self> {
        let limits = options.source_limits_value();
        if bytes.len() > limits.max_package_bytes {
            return Err(PackageError::Limit {
                kind: crate::SourceLimitKind::InputBytes,
                observed: bytes.len(),
                maximum: limits.max_package_bytes,
            }
            .into());
        }
        let mut package = Reader::from_reader_with_options(Cursor::new(bytes.as_slice()), options)?;
        let has_signature = package.contains_signature_component()?;
        let presentation = package.presentation_with_options(options)?;
        drop(package);
        let artifact = Arc::from(bytes.into_boxed_slice());
        Self::from_admitted(artifact, presentation, has_signature, options)
    }

    /// Capture a native snapshot from an already admitted canonical package.
    ///
    /// The package has already validated and retained its CFB index and
    /// presentation projection. Reusing that projection avoids opening the
    /// same source container a second time when a package-level edit starts;
    /// the exact bytes and the source freshness fence remain explicit.
    pub(crate) fn from_package_with_options(
        package: &crate::Package,
        bytes: Arc<[u8]>,
        options: OpenOptions<'_>,
    ) -> Result<Self> {
        #[cfg(feature = "vba-inspection")]
        let options = options.inherit_vba_project_limits(package.vba_project_limits());
        let presentation = package.presentation()?;
        let has_signature = package
            .exact()
            .signed(|| package.contains_signature_component())?;
        Self::from_admitted(bytes, presentation, has_signature, options)
    }

    fn from_admitted(
        artifact: Arc<[u8]>,
        presentation: crate::presentation::Presentation,
        has_signature: bool,
        options: OpenOptions<'_>,
    ) -> Result<Self> {
        presentation.check_source()?;
        let limits = options.source_limits_value();
        if artifact.len() > limits.max_package_bytes {
            return Err(PackageError::Limit {
                kind: crate::SourceLimitKind::InputBytes,
                observed: artifact.len(),
                maximum: limits.max_package_bytes,
            }
            .into());
        }
        let directory = presentation.slide_directory();
        // Resolve the live document through the already parsed presentation.
        // The embedded append-only editor intentionally walks UserEdit chains
        // strictly; using it here would discard a caller-selected tolerant
        // read before we can record whether that repair is editable.
        let document_persist_id = directory.document_persist_id();
        let document_offset = directory.document_offset();
        let document_stream = presentation.document_stream();
        let (document_record, document_consumed) =
            crate::Record::parse(document_stream, document_offset, &limits)?;
        if document_record.record_type != crate::RecordType::Document {
            return Err(PackageError::Corrupted(
                "PPT slide-order snapshot resolved a non-document live record".into(),
            )
            .into());
        }
        let document_end = document_offset
            .checked_add(document_consumed)
            .ok_or_else(|| {
                PackageError::Corrupted("PPT live document record offset overflows".into())
            })?;
        let document_bytes = document_stream
            .get(document_offset..document_end)
            .ok_or_else(|| {
                PackageError::Corrupted("PPT live document record exceeds its stream".into())
            })?
            .to_vec();
        let document_limits = document_structure::Limits {
            max_bytes: limits.max_record_bytes.min(limits.max_input_bytes),
            max_records: limits.max_records,
            max_depth: limits.max_depth,
        };
        let document =
            document_structure::Snapshot::parse_with_limits(document_bytes, document_limits)?;
        if document.slides().len() != directory.len()
            || !document
                .slides()
                .iter()
                .zip(directory.entries())
                .all(|(slide, entry)| {
                    slide.slide_id() == entry.slide_id() && slide.persist_id() == entry.persist_id()
                })
        {
            return Err(PackageError::Corrupted(
                "PPT live document and slide directory orders disagree".into(),
            )
            .into());
        }
        let review = presentation.document_comparison()?.review()?;
        let has_review_history = review.diff_tree_count() != 0;
        let chain_repaired = presentation
            .tolerance_report()
            .defects()
            .iter()
            .any(|defect| matches!(defect.defect, crate::leniency::Defect::UserEditChain(_)))
            || presentation.tolerance_report().unrecorded() != 0;
        presentation.check_source()?;
        Ok(Self {
            artifact: Artifact::new(artifact),
            document,
            document_persist_id,
            limits,
            leniency: options.leniency_value(),
            #[cfg(feature = "vba-inspection")]
            vba_project_limits: options.vba_project_limits_value(),
            chain_repaired,
            has_review_history,
            has_signature,
        })
    }

    fn adopt_text_edit_publication(
        &self,
        publication: crate::text_edit::RootPublication,
        expected_slide_persist_id: u32,
    ) -> Result<Self> {
        if publication.source() != self.bytes() {
            return Err(PackageError::Corrupted(
                "PPT root text publication source changed before adoption".into(),
            )
            .into());
        }
        let replaced_slide_persist_id = publication.replaced_slide_persist_id();
        if replaced_slide_persist_id == self.document_persist_id
            || expected_slide_persist_id == self.document_persist_id
        {
            return Err(PackageError::Corrupted(
                "PPT root text publication targeted the live document record".into(),
            )
            .into());
        }
        if replaced_slide_persist_id != expected_slide_persist_id {
            return Err(PackageError::Corrupted(
                "PPT root text publication targeted an unexpected slide record".into(),
            )
            .into());
        }

        let bytes = publication.into_output();
        if self.limits != SourceLimits::default() {
            return Self::from_bytes_with_options(bytes.to_vec(), self.open_options());
        }
        if bytes.len() > self.limits.max_package_bytes {
            return Err(PackageError::Limit {
                kind: crate::SourceLimitKind::InputBytes,
                observed: bytes.len(),
                maximum: self.limits.max_package_bytes,
            }
            .into());
        }
        Ok(Self {
            artifact: Artifact::new(bytes),
            document: self.document.clone(),
            document_persist_id: self.document_persist_id,
            limits: self.limits,
            leniency: self.leniency,
            #[cfg(feature = "vba-inspection")]
            vba_project_limits: self.vba_project_limits,
            chain_repaired: self.chain_repaired,
            has_review_history: self.has_review_history,
            has_signature: self.has_signature,
        })
    }

    fn adopt_notes_publication(
        &self,
        publication: crate::text_edit::RootNotesPublication,
    ) -> Result<Self> {
        if publication.source() != self.bytes() {
            return Err(PackageError::Corrupted(
                "PPT notes text publication source changed before adoption".into(),
            )
            .into());
        }
        if publication.notes_persist_id() == self.document_persist_id {
            return Err(PackageError::Corrupted(
                "PPT notes text publication targeted the live document record".into(),
            )
            .into());
        }
        let bytes = publication.into_output();
        if self.limits != SourceLimits::default() {
            return Self::from_bytes_with_options(bytes.to_vec(), self.open_options());
        }
        if bytes.len() > self.limits.max_package_bytes {
            return Err(PackageError::Limit {
                kind: crate::SourceLimitKind::InputBytes,
                observed: bytes.len(),
                maximum: self.limits.max_package_bytes,
            }
            .into());
        }
        Ok(Self {
            artifact: Artifact::new(bytes),
            document: self.document.clone(),
            document_persist_id: self.document_persist_id,
            limits: self.limits,
            leniency: self.leniency,
            #[cfg(feature = "vba-inspection")]
            vba_project_limits: self.vba_project_limits,
            chain_repaired: self.chain_repaired,
            has_review_history: self.has_review_history,
            has_signature: self.has_signature,
        })
    }

    fn adopt_text_edit_batch_publication(
        &self,
        publication: crate::text_edit::RootBatchPublication,
    ) -> Result<(Self, Vec<crate::text_edit::RootShapeTextChange>)> {
        if publication.source() != self.bytes() {
            return Err(PackageError::Corrupted(
                "PPT root batch text publication source changed before adoption".into(),
            )
            .into());
        }
        if publication.changes().is_empty() {
            return Err(PackageError::Corrupted(
                "changed PPT batch text publication has no semantic changes".into(),
            )
            .into());
        }
        if publication
            .changes()
            .iter()
            .any(|change| change.slide_persist_id() == self.document_persist_id)
        {
            return Err(PackageError::Corrupted(
                "PPT root batch text publication targeted the live document record".into(),
            )
            .into());
        }

        let document_changed = publication.document_changed();
        let (bytes, changes) = publication.into_parts();
        if document_changed {
            // Outline-backed text lives in the live document: reopen so the
            // retained document structure is the published one.
            let reopened = Self::from_bytes_with_options(bytes.to_vec(), self.open_options())?;
            if reopened.document.slides() != self.document.slides() {
                return Err(PackageError::Corrupted(
                    "PPT outline text publication changed slide structure".into(),
                )
                .into());
            }
            return Ok((reopened, changes));
        }
        if self.limits != SourceLimits::default() {
            let reopened = Self::from_bytes_with_options(bytes.to_vec(), self.open_options())?;
            return Ok((reopened, changes));
        }
        if bytes.len() > self.limits.max_package_bytes {
            return Err(PackageError::Limit {
                kind: crate::SourceLimitKind::InputBytes,
                observed: bytes.len(),
                maximum: self.limits.max_package_bytes,
            }
            .into());
        }
        Ok((
            Self {
                artifact: Artifact::new(bytes),
                document: self.document.clone(),
                document_persist_id: self.document_persist_id,
                limits: self.limits,
                leniency: self.leniency,
                #[cfg(feature = "vba-inspection")]
                vba_project_limits: self.vba_project_limits,
                chain_repaired: self.chain_repaired,
                has_review_history: self.has_review_history,
                has_signature: self.has_signature,
            },
            changes,
        ))
    }

    /// Exact bytes of the complete source or committed artifact.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.artifact.bytes()
    }

    /// Number of presentation slides in semantic order.
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.document.slides().len()
    }

    /// Whether one slide is hidden during the slide show.
    ///
    /// A slide without a `SlideShowSlideInfoAtom` is shown (`false`).
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the position is absent or the slide has
    /// several or a malformed slide-show information atom.
    pub fn slide_hidden(&self, position: Position) -> Result<bool> {
        self.require_position(position)?;
        inspect_slide_hidden(self, position)
    }

    /// Returns one slide's fixed-width manual/automatic advance state.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the position is absent or does not own
    /// exactly one valid `SlideShowSlideInfoAtom`.
    pub fn slide_advance(&self, position: Position) -> Result<SlideAdvance> {
        self.require_position(position)?;
        inspect_slide_advance(self, position)
    }

    /// Returns one slide's fixed-width visual transition state.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the position is absent or does not own one
    /// canonical `SlideShowSlideInfoAtom` visual mapping.
    pub fn slide_transition_visual(&self, position: Position) -> Result<SlideTransitionVisual> {
        self.require_position(position)?;
        inspect_slide_transition_visual(self, position)
    }

    /// Returns the checked complete host anchor for one existing shape.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the target is absent, ambiguous, or does
    /// not own exactly one supported `PowerPoint` client anchor.
    pub fn shape_anchor(&self, target: crate::text_edit::Target) -> Result<crate::Anchor> {
        self.require_position(target.slide())?;
        if self.has_signature {
            return crate::text_edit::inspect_shape_anchor_readonly(
                self.bytes(),
                target,
                self.open_options(),
            )
            .map_err(Error::from);
        }
        crate::text_edit::inspect_shape_anchor(self.bytes(), target).map_err(Error::from)
    }

    /// Returns the inert path of one linked movie/audio relationship.
    ///
    /// # Errors
    ///
    /// Returns an error when the live document's media owner is malformed, the
    /// identifier is absent, or it addresses a non-path media kind.
    pub fn external_media_path(&self, id: u32) -> Result<Option<String>> {
        let media = external_media_snapshot(self)?;
        let object = media
            .collection()
            .and_then(|collection| collection.objects.iter().find(|object| object.id() == id))
            .ok_or_else(|| {
                PackageError::InvalidFormat(format!("external media ID {id} was not found"))
            })?;
        media_path_of(object)
    }

    /// Returns the playback flags of one external-media relationship.
    ///
    /// # Errors
    ///
    /// Returns an error when the live document's media owner is malformed or
    /// the identifier is absent.
    pub fn external_media_playback(&self, id: u32) -> Result<crate::external_media::Playback> {
        external_media_snapshot(self)?
            .collection()
            .and_then(|collection| collection.playback(id))
            .ok_or_else(|| {
                PackageError::InvalidFormat(format!("external media ID {id} was not found")).into()
            })
    }

    /// Resource bounds retained for commits and patch application.
    #[must_use]
    pub(crate) const fn limits(&self) -> SourceLimits {
        self.limits
    }

    /// Leniency policy retained by this immutable snapshot.
    #[must_use]
    pub const fn leniency(&self) -> crate::leniency::Leniency {
        self.leniency
    }

    fn open_options(&self) -> OpenOptions<'static> {
        let options = OpenOptions::new()
            .source_limits(self.limits)
            .leniency(self.leniency);
        #[cfg(feature = "vba-inspection")]
        let options = options.vba_project_limits(self.vba_project_limits);
        options
    }

    /// Starts an isolated source-bound order edit.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for encrypted or metadata-bearing CFB sources
    /// and for position-dependent reviewer history. Signed sources are
    /// admitted for exact no-op planning; changed operations return
    /// [`Refusal::Signed`] before publication.
    pub(crate) fn edit(&self) -> Result<Transaction> {
        self.require_editable()?;
        Ok(Transaction {
            source: self.clone(),
            working: self.clone(),
            document: self.document.edit(),
            changes: Vec::new(),
            structural: Vec::new(),
            visibility_changes: Vec::new(),
            advance_changes: Vec::new(),
            transition_visual_changes: Vec::new(),
            text_changes: Vec::new(),
            anchor_changes: Vec::new(),
            media_path_changes: Vec::new(),
            media_playback_changes: Vec::new(),
            formatting: Vec::new(),
            inserted_records: BTreeMap::new(),
            inserted_pictures: BTreeMap::new(),
            live_document_changed: false,
            opened: None,
            verified_shapes: None,
        })
    }

    /// Plans a dependency-checked transfer from `donor` without mutating
    /// either presentation.
    ///
    /// The bounded closure reuses a byte-identical master plus matching
    /// comment-author, hyperlink, sound, and external-media owners already
    /// present in the receiver. Ordinary shapes and their passive reference
    /// rules are deterministically rewritten to target-owned unused drawing
    /// IDs. Referenced BLIPs are reused by semantic identity or appended with
    /// exact native image framing to the target `BStore` and `Pictures` stream.
    /// Notes, animation/build graphs, active OLE/macro/program actions, and
    /// missing or mismatched remaining owners are refused explicitly.
    ///
    /// # Errors
    ///
    /// Returns a typed dependency refusal, a position refusal, or a package
    /// error for malformed source records.
    pub(crate) fn plan_transfer_from(&self, donor: &Self, slide: Position) -> Result<TransferPlan> {
        self.require_editable()?;
        if self.has_signature || donor.has_signature {
            return Err(Error::Refused(Refusal::Signed));
        }
        donor.require_position(slide)?;
        let donor_slide = donor.document.slides()[slide.get()];
        let group = donor.document.edit().slide_group(slide.get())?;
        let mut record = persisted_record(donor, donor_slide.persist_id())?;
        let closure = require_portable_slide(self, donor, &mut record)?;
        let payload = slide_payload(group, record)?;
        require_master_closure(self, donor, payload.master_id)?;
        Ok(TransferPlan {
            target_artifact: self.artifact_digest().to_owned(),
            payload: normalized_payload(&SlidePayload {
                pictures: closure.pictures,
                ..payload
            }),
            reused_dependencies: closure.dependencies,
            relationship_remaps: closure.remaps,
        })
    }

    /// Produces a non-mutating three-way conflict plan for two patches based
    /// on this exact snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when either patch was not prepared from this base.
    pub(crate) fn plan_three_way(&self, left: &Patch, right: &Patch) -> Result<ThreeWayPlan> {
        if left.before() != self.bytes() || right.before() != self.bytes() {
            return Err(PackageError::InvalidFormat(
                "PPT three-way planning requires two patches from the exact base artifact".into(),
            )
            .into());
        }
        Ok(ThreeWayPlan::new(left, right))
    }

    /// Applies supported durable slide-order operations to this exact source.
    ///
    /// # Errors
    ///
    /// Returns an error for an unsupported vocabulary, malformed selector,
    /// failed exact-artifact/order precondition, or refused move.
    pub(crate) fn apply_durable<Mode>(
        &self,
        patch: &litchi_core::patch::Patch<Mode>,
    ) -> Result<Self> {
        if patch.format() != "litchi-ppt" {
            return Err(invalid_durable_patch("unsupported durable patch format"));
        }
        if patch.operations().is_empty() {
            return Ok(self.clone());
        }
        let mut current = self.clone();
        let operations = patch.operations();
        let mut index = 0usize;
        while index < operations.len() {
            if is_formatting_operation(&operations[index].op) {
                current = match operations[index].op.as_str() {
                    "presentation.shape-text.set" => {
                        apply_durable_text(&current, &operations[index])?
                    },
                    "presentation.notes-text.set" => {
                        apply_durable_notes_text(&current, &operations[index])?
                    },
                    "presentation.shape-anchor.set" => {
                        apply_durable_anchor(&current, &operations[index])?
                    },
                    "presentation.slide-hidden.set" => {
                        apply_durable_slide_visibility(&current, &operations[index])?
                    },
                    "presentation.slide-advance.set" => {
                        apply_durable_slide_advance(&current, &operations[index])?
                    },
                    "presentation.slide-transition-visual.set" => {
                        apply_durable_slide_transition_visual(&current, &operations[index])?
                    },
                    "presentation.external-media-path.set" => {
                        apply_durable_media_path(&current, &operations[index])?
                    },
                    "presentation.external-media-playback.set" => {
                        apply_durable_media_playback(&current, &operations[index])?
                    },
                    "presentation.element-records.set" => {
                        records::apply_durable(&current, &operations[index], patch.blobs())?
                    },
                    _ => unreachable!("formatting vocabulary is checked by its predicate"),
                };
                index += 1;
                continue;
            }

            // The transaction's source is a clone of `current`, so its
            // insertion target check reuses this memoized digest, and the
            // inner commit computes none.
            let structural_artifact = current.artifact_digest().to_owned();
            let mut transaction = current.edit()?;
            while index < operations.len() && !is_formatting_operation(&operations[index].op) {
                apply_durable_structural(
                    &mut transaction,
                    &operations[index],
                    patch.blobs(),
                    &structural_artifact,
                )?;
                index += 1;
            }
            current = transaction.commit()?.snapshot;
        }
        Ok(current)
    }

    fn require_editable(&self) -> Result<()> {
        // A signed source may still be opened for a read-only transaction.
        // The individual operation owners perform the changed/no-op split so
        // that an exact no-op can retain the original CFB, including nested
        // signature storages, without invoking the mutation editor.  This
        // admission also deliberately precedes history-closure refusals:
        // returning the exact source does not depend on appending a new edit.
        if self.has_signature {
            return Ok(());
        }
        // The append-only native editor intentionally reads the UserEdit
        // chain with strict semantics.  A tolerated cyclic chain is safe for
        // read projection, but cannot yet be used as an edit base because
        // appending a new edit would otherwise retain an ambiguous history.
        // Clean packages opened under the same leniency policy remain
        // editable; selecting a policy alone is not a refusal.
        if self.chain_repaired {
            return Err(Error::Refused(Refusal::RepairedUserEditChain));
        }
        if self.has_review_history {
            return Err(Error::Refused(Refusal::ReviewHistoryDependency));
        }
        crate::font::require_stream_only_cfb(self.bytes())
            .and_then(|()| {
                crate::embedded::object::Editor::open_records_arc_with_limit(
                    self.shared_bytes(),
                    self.limits.max_package_bytes,
                    self.limits,
                )
                .map(|_editor| ())
            })
            .map_err(|_error| Error::Refused(Refusal::UnsupportedSource))
    }

    fn require_position(&self, position: Position) -> Result<()> {
        if position.get() >= self.slide_count() {
            return Err(Error::Refused(Refusal::SlideNotFound { position }));
        }
        Ok(())
    }

    /// Lowercase hexadecimal SHA-256 of the exact artifact, computed on first
    /// use and shared by every clone of this snapshot.
    fn artifact_digest(&self) -> &str {
        self.artifact.sha256()
    }

    /// Shared handle to the exact artifact bytes.
    fn shared_bytes(&self) -> Arc<[u8]> {
        Arc::clone(self.artifact.shared())
    }
}

/// One semantic move staged in a transaction or durable patch.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Change {
    from: Position,
    destination: Position,
    before_order: [u8; 32],
    after_order: [u8; 32],
}

impl Change {
    /// Source position resolved against the transaction-local order.
    #[must_use]
    pub(crate) const fn from(&self) -> Position {
        self.from
    }

    /// Final destination position for the selected slide.
    #[must_use]
    pub(crate) const fn destination(&self) -> Position {
        self.destination
    }

    fn inverse(self) -> Self {
        Self {
            from: self.destination,
            destination: self.from,
            before_order: self.after_order,
            after_order: self.before_order,
        }
    }
}

/// One checked text replacement preserving the selected shape's formatting
/// and modeled dependency closure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ShapeTextChange {
    target: crate::text_edit::Target,
    before: String,
    after: String,
}

/// One checked speaker-notes text replacement on the slide at a source
/// position.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct NotesTextChange {
    slide: Position,
    before: String,
    after: String,
}

impl NotesTextChange {
    /// The slide whose notes changed, in the source slide order.
    #[must_use]
    pub(crate) const fn slide(&self) -> Position {
        self.slide
    }

    /// The notes text before the change.
    #[must_use]
    pub(crate) fn before(&self) -> &str {
        &self.before
    }

    /// The notes text after the change.
    #[must_use]
    pub(crate) fn after(&self) -> &str {
        &self.after
    }
}

/// One contextual shape-text replacement for
/// [`Edit::set_shape_texts`](crate::package::edit::Edit::set_shape_texts).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeTextReplacement {
    target: crate::text_edit::Target,
    text: String,
}

impl ShapeTextReplacement {
    /// Creates a replacement for one shape in the transaction's projected
    /// slide order.
    #[must_use]
    pub fn new(target: crate::text_edit::Target, text: impl Into<String>) -> Self {
        Self {
            target,
            text: text.into(),
        }
    }

    /// Semantic target in the transaction's projected slide order.
    #[must_use]
    pub const fn target(&self) -> crate::text_edit::Target {
        self.target
    }

    /// Replacement text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// Flattens an aggregate shape-text plan only after all finite admission
/// dimensions have been checked.
///
/// The returned descriptors borrow the caller's replacement strings. Keeping
/// this admission separate from target resolution means an oversized or
/// malformed chunk plan cannot trigger package parsing or proportional native
/// work. The byte count is checked here as well as in the native publisher so
/// a future caller cannot accidentally turn an aggregate sum into a wrapping
/// counter.
fn flatten_shape_text_chunks<'a>(
    chunks: &[&'a [ShapeTextReplacement]],
) -> Result<(Vec<(crate::text_edit::Target, &'a str)>, usize)> {
    if chunks.len() > MAX_SHAPE_TEXT_CHUNKS {
        return Err(PackageError::ResourceLimit(format!(
            "PPT shape-text chunk count {} exceeds limit {}",
            chunks.len(),
            MAX_SHAPE_TEXT_CHUNKS
        ))
        .into());
    }

    let mut selection_count = 0usize;
    let mut replacement_bytes = 0usize;
    for chunk in chunks {
        if chunk.len() > MAX_SHAPE_TEXT_REPLACEMENTS {
            return Err(PackageError::ResourceLimit(format!(
                "PPT shape-text chunk operation count {} exceeds limit {}",
                chunk.len(),
                MAX_SHAPE_TEXT_REPLACEMENTS
            ))
            .into());
        }
        selection_count = selection_count.checked_add(chunk.len()).ok_or_else(|| {
            PackageError::ResourceLimit("PPT shape-text aggregate operation count overflow".into())
        })?;
        if selection_count > MAX_SHAPE_TEXT_CHUNK_SELECTIONS {
            return Err(PackageError::ResourceLimit(format!(
                "PPT shape-text aggregate operation count {} exceeds limit {}",
                selection_count, MAX_SHAPE_TEXT_CHUNK_SELECTIONS
            ))
            .into());
        }
        for replacement in *chunk {
            replacement_bytes = replacement_bytes
                .checked_add(replacement.text().len())
                .ok_or_else(|| {
                    PackageError::ResourceLimit(
                        "PPT shape-text aggregate replacement byte count overflow".into(),
                    )
                })?;
        }
    }

    let mut flattened = Vec::with_capacity(selection_count);
    for chunk in chunks {
        flattened.extend(
            chunk
                .iter()
                .map(|replacement| (replacement.target(), replacement.text())),
        );
    }
    Ok((flattened, replacement_bytes))
}

impl ShapeTextChange {
    /// Semantic source-order shape target.
    #[must_use]
    pub(crate) const fn target(&self) -> crate::text_edit::Target {
        self.target
    }

    /// Text required before the replacement.
    #[must_use]
    pub(crate) fn before(&self) -> &str {
        &self.before
    }

    /// Replacement text.
    #[must_use]
    pub(crate) fn after(&self) -> &str {
        &self.after
    }
}

/// One checked shape-geometry formatting replacement.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct ShapeAnchorChange {
    target: crate::text_edit::Target,
    before: crate::Anchor,
    after: crate::Anchor,
}

/// One checked replacement of an inert external movie/audio path.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ExternalMediaPathChange {
    id: u32,
    before: Option<String>,
    after: Option<String>,
}

impl ExternalMediaPathChange {
    /// External-object identifier retained by the replacement.
    #[must_use]
    pub(crate) const fn id(&self) -> u32 {
        self.id
    }

    /// Path required before the replacement.
    #[must_use]
    pub(crate) fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    /// Replacement path, or `None` to clear it.
    #[must_use]
    pub(crate) fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }
}

/// One checked replacement of external-media playback flags.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct ExternalMediaPlaybackChange {
    id: u32,
    before: crate::external_media::Playback,
    after: crate::external_media::Playback,
}

/// One fixed-width replacement of a slide's hidden flag.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct SlideVisibilityChange {
    position: Position,
    target: Position,
    slide_id: u32,
    before: bool,
    after: bool,
}

impl SlideVisibilityChange {
    /// Semantic slide position resolved against the transaction-local order.
    #[must_use]
    pub(crate) const fn position(self) -> Position {
        self.position
    }

    /// Hidden state required before the replacement.
    #[must_use]
    pub(crate) const fn before(self) -> bool {
        self.before
    }

    /// Replacement hidden state.
    #[must_use]
    pub(crate) const fn after(self) -> bool {
        self.after
    }
}

/// One fixed-width replacement of slide-show advance timing and flags.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct SlideAdvanceChange {
    target: Position,
    slide_id: u32,
    before: SlideAdvance,
    after: SlideAdvance,
}

impl SlideAdvanceChange {
    /// Semantic source-order slide target.
    #[must_use]
    pub(crate) const fn target(self) -> Position {
        self.target
    }

    /// Advance state required before replacement.
    #[must_use]
    pub(crate) const fn before(self) -> SlideAdvance {
        self.before
    }

    /// Replacement advance state.
    #[must_use]
    pub(crate) const fn after(self) -> SlideAdvance {
        self.after
    }
}

/// One fixed-width replacement of a slide's visual transition fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct SlideTransitionVisualChange {
    target: Position,
    slide_id: u32,
    before: SlideTransitionVisual,
    after: SlideTransitionVisual,
}

impl SlideTransitionVisualChange {
    /// Semantic source-order slide target.
    #[must_use]
    pub(crate) const fn target(self) -> Position {
        self.target
    }

    /// Visual transition required before replacement.
    #[must_use]
    pub(crate) const fn before(self) -> SlideTransitionVisual {
        self.before
    }

    /// Replacement visual transition.
    #[must_use]
    pub(crate) const fn after(self) -> SlideTransitionVisual {
        self.after
    }
}

impl ExternalMediaPlaybackChange {
    /// External-object identifier retained by the replacement.
    #[must_use]
    pub(crate) const fn id(self) -> u32 {
        self.id
    }

    /// Playback flags required before the replacement.
    #[must_use]
    pub(crate) const fn before(self) -> crate::external_media::Playback {
        self.before
    }

    /// Replacement playback flags.
    #[must_use]
    pub(crate) const fn after(self) -> crate::external_media::Playback {
        self.after
    }
}

impl ShapeAnchorChange {
    /// Semantic source-order shape target.
    #[must_use]
    pub(crate) const fn target(self) -> crate::text_edit::Target {
        self.target
    }

    /// Anchor required before the replacement.
    #[must_use]
    pub(crate) const fn before(self) -> crate::Anchor {
        self.before
    }

    /// Replacement anchor.
    #[must_use]
    pub(crate) const fn after(self) -> crate::Anchor {
        self.after
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum FormattingChange {
    Text(ShapeTextChange),
    NotesText(NotesTextChange),
    Anchor(ShapeAnchorChange),
    SlideVisibility(SlideVisibilityChange),
    SlideAdvance(SlideAdvanceChange),
    SlideTransitionVisual(SlideTransitionVisualChange),
    MediaPath(ExternalMediaPathChange),
    MediaPlayback(ExternalMediaPlaybackChange),
    /// Persisted records an element edit replaced (record 0323).
    Element(records::ElementChange),
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct SlidePayload {
    group: Vec<crate::records::Record>,
    record: Vec<u8>,
    master_id: u32,
    pictures: Vec<TransferredPicture>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct TransferredPicture {
    target_id: u32,
    blip: Vec<u8>,
}

/// Non-mutating dependency-checked plan for transferring one slide from a
/// donor presentation into one exact receiving presentation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TransferPlan {
    target_artifact: String,
    payload: SlidePayload,
    reused_dependencies: Vec<TransferDependency>,
    relationship_remaps: Vec<RelationshipRemap>,
}

/// One donor-native relationship ID rewritten to a semantically equivalent
/// target-native relationship during bounded transfer planning.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RelationshipRemap {
    dependency: TransferDependency,
    source_id: u32,
    target_id: u32,
}

impl RelationshipRemap {
    /// Owner family whose identifier is remapped.
    #[must_use]
    pub const fn dependency(self) -> TransferDependency {
        self.dependency
    }

    /// Native donor identifier found in the slide.
    #[must_use]
    pub const fn source_id(self) -> u32 {
        self.source_id
    }

    /// Semantically equivalent native identifier in the target.
    #[must_use]
    pub const fn target_id(self) -> u32 {
        self.target_id
    }
}

impl TransferPlan {
    /// Master identity that must already be present in the receiving deck.
    #[must_use]
    pub const fn required_master_id(&self) -> u32 {
        self.payload.master_id
    }

    /// Presentation-global owners proven identical and therefore reused by
    /// this bounded transfer without rewriting their native identifiers.
    #[must_use]
    pub fn reused_dependencies(&self) -> &[TransferDependency] {
        &self.reused_dependencies
    }

    /// Presentation-global owners resolved by semantic reuse or fixed-width
    /// native-ID remapping for this plan.
    #[must_use]
    pub fn resolved_dependencies(&self) -> &[TransferDependency] {
        &self.reused_dependencies
    }

    /// Native relationship rewrites already applied to the staged slide
    /// payload. Identity mappings are omitted.
    #[must_use]
    pub fn relationship_remaps(&self) -> &[RelationshipRemap] {
        &self.relationship_remaps
    }
}

/// One deterministic semantic overlap found during non-mutating three-way
/// planning.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ThreeWayConflict {
    target: String,
}

impl ThreeWayConflict {
    /// Stable semantic target shared by both branches.
    #[must_use]
    pub(crate) fn target(&self) -> &str {
        &self.target
    }
}

/// Read-only conflict result for two patches prepared from one exact base.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ThreeWayPlan {
    conflicts: Vec<ThreeWayConflict>,
}

impl ThreeWayPlan {
    pub(crate) fn new(left: &Patch, right: &Patch) -> Self {
        let left_effects = patch_effects(left);
        let right_effects = patch_effects(right);
        let conflicts = left_effects
            .iter()
            .filter(|target| right_effects.contains(*target))
            .map(|target| ThreeWayConflict {
                target: target.clone(),
            })
            .collect();
        Self { conflicts }
    }

    /// Whether the two branches have no overlapping semantic write effects.
    #[must_use]
    pub(crate) fn is_clean(&self) -> bool {
        self.conflicts.is_empty()
    }

    /// Stable sorted conflict set.
    #[must_use]
    pub(crate) fn conflicts(&self) -> &[ThreeWayConflict] {
        &self.conflicts
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ListChange {
    position: Position,
    before_order: [u8; 32],
    after_order: [u8; 32],
    payload: SlidePayload,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum StructuralChange {
    Move(Change),
    Remove(ListChange),
    Insert(ListChange),
    /// A `DocumentContainer` edit outside the slide list (record 0357a),
    /// with the slide-order fingerprint it was staged at (record 0357b:
    /// the edit leaves the slide order alone, so the fingerprint holds on
    /// both sides and binds its durable operation in either direction).
    Document(DocumentChange, [u8; 32]),
}

/// A staged `DocumentContainer` edit that leaves the slide list order alone:
/// master order, the custom table-style atom and `fNonOutlineData`
/// (record 0357a).
#[derive(Debug, Clone, PartialEq, Eq)]
enum DocumentChange {
    /// The master list became `order` (each entry a previous position).
    Masters { order: Vec<usize> },
    /// A custom table-style atom was inserted.
    InsertStyles {
        record: crate::Record,
        placement: document_structure::CustomTableStylesPlacement,
    },
    /// The custom table-style atom was removed.
    RemoveStyles {
        record: crate::Record,
        placement: document_structure::CustomTableStylesPlacement,
    },
    /// The custom table-style atom was replaced in place.
    ReplaceStyles {
        before: crate::Record,
        after: crate::Record,
    },
    /// The custom table-style atom moved around the `EndDocumentAtom`.
    MoveStyles {
        before: document_structure::CustomTableStylesPlacement,
        after: document_structure::CustomTableStylesPlacement,
    },
    /// One slide's `SlidePersistAtom.fNonOutlineData` changed.
    NonOutlineData {
        position: usize,
        before: bool,
        after: bool,
    },
}

impl DocumentChange {
    fn inverse(&self) -> Self {
        match self {
            Self::Masters { order } => {
                let mut inverse = vec![0; order.len()];
                for (position, previous) in order.iter().enumerate() {
                    inverse[*previous] = position;
                }
                Self::Masters { order: inverse }
            },
            Self::InsertStyles { record, placement } => Self::RemoveStyles {
                record: record.clone(),
                placement: *placement,
            },
            Self::RemoveStyles { record, placement } => Self::InsertStyles {
                record: record.clone(),
                placement: *placement,
            },
            Self::ReplaceStyles { before, after } => Self::ReplaceStyles {
                before: after.clone(),
                after: before.clone(),
            },
            Self::MoveStyles { before, after } => Self::MoveStyles {
                before: *after,
                after: *before,
            },
            Self::NonOutlineData {
                position,
                before,
                after,
            } => Self::NonOutlineData {
                position: *position,
                before: *after,
                after: *before,
            },
        }
    }

    fn replay(&self, document: &mut document_structure::Transaction) -> Result<()> {
        match self {
            Self::Masters { order } => document.reorder_masters(order)?,
            Self::InsertStyles { record, placement } => {
                document.insert_custom_table_styles(record.clone(), *placement)?;
            },
            Self::RemoveStyles { .. } => {
                document.remove_custom_table_styles()?;
            },
            Self::ReplaceStyles { after, .. } => {
                document.replace_custom_table_styles(after.clone())?;
            },
            Self::MoveStyles { after, .. } => document.move_custom_table_styles(*after)?,
            Self::NonOutlineData {
                position, after, ..
            } => document.set_slide_non_outline_data(*position, *after)?,
        }
        Ok(())
    }

    /// The semantic target this change writes, for three-way planning.
    fn effect(&self) -> String {
        match self {
            Self::Masters { .. } => "document/master-list".to_owned(),
            Self::InsertStyles { .. }
            | Self::RemoveStyles { .. }
            | Self::ReplaceStyles { .. }
            | Self::MoveStyles { .. } => "document/custom-table-styles".to_owned(),
            Self::NonOutlineData { position, .. } => {
                format!("slide-list/position:{position}/non-outline-data")
            },
        }
    }
}

/// Isolated failure-atomic edit over one opened presentation's slide order.
#[derive(Debug, Clone)]
pub(crate) struct Transaction {
    source: Snapshot,
    working: Snapshot,
    document: document_structure::Transaction,
    changes: Vec<Change>,
    structural: Vec<StructuralChange>,
    visibility_changes: Vec<SlideVisibilityChange>,
    advance_changes: Vec<SlideAdvanceChange>,
    transition_visual_changes: Vec<SlideTransitionVisualChange>,
    text_changes: Vec<ShapeTextChange>,
    anchor_changes: Vec<ShapeAnchorChange>,
    media_path_changes: Vec<ExternalMediaPathChange>,
    media_playback_changes: Vec<ExternalMediaPlaybackChange>,
    formatting: Vec<FormattingChange>,
    inserted_records: BTreeMap<u32, Vec<u8>>,
    inserted_pictures: BTreeMap<u32, Vec<u8>>,
    /// Whether a staged text edit republished the live document (outline
    /// text), so staged structure must be rebased over it at commit.
    live_document_changed: bool,
    /// The opened projection of exactly the source bytes, when the root
    /// edit was started from an immutable package (record 0260).
    opened: Option<OpenedProjection>,
    /// Shapes of one slide of the working artifact, proven equal to a parse
    /// by the last batch text publication's incremental readback (record
    /// 0264).
    verified_shapes: Option<VerifiedShapes>,
}

/// Shapes a parse of slide `persist_id` of exactly `bytes` returns.
pub(crate) struct VerifiedSlideShapes {
    /// The exact artifact the shapes describe.
    pub(crate) bytes: Arc<[u8]>,
    /// The slide's persist identifier in that artifact.
    pub(crate) persist_id: u32,
    /// What parsing that slide returns.
    pub(crate) shapes: Arc<crate::slide::PatchedShapes>,
}

/// Shapes a parse of slide `persist_id` of exactly `bytes` returns.
#[derive(Clone)]
struct VerifiedShapes {
    bytes: Arc<[u8]>,
    persist_id: u32,
    shapes: Arc<crate::slide::PatchedShapes>,
}

impl fmt::Debug for VerifiedShapes {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("VerifiedShapes")
            .field("persist_id", &self.persist_id)
            .finish_non_exhaustive()
    }
}

/// An opened projection of a transaction's exact source bytes. It is used
/// only while the working artifact is still the source allocation, so a
/// target resolves against the parse the caller already holds.
#[derive(Clone)]
struct OpenedProjection(crate::Presentation);

impl fmt::Debug for OpenedProjection {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("OpenedProjection")
    }
}

impl Transaction {
    /// Attach the opened projection of exactly this transaction's source
    /// bytes, which target resolution reuses while nothing has been staged.
    pub(crate) fn with_opened_projection(mut self, presentation: crate::Presentation) -> Self {
        self.opened = Some(OpenedProjection(presentation));
        self
    }

    /// The opened source projection, while the working artifact is still the
    /// exact source allocation.
    fn opened_source(&self) -> Option<&crate::Presentation> {
        self.opened
            .as_ref()
            .filter(|_| Arc::ptr_eq(&self.working.shared_bytes(), &self.source.shared_bytes()))
            .map(|opened| &opened.0)
    }

    /// The working artifact and the proven shapes of one of its slides, when
    /// the last batch text publication verified exactly this working
    /// artifact (record 0264).
    pub(crate) fn verified_shapes(&self) -> Option<VerifiedSlideShapes> {
        self.verified_shapes
            .as_ref()
            .filter(|verified| Arc::ptr_eq(&verified.bytes, &self.working.shared_bytes()))
            .map(|verified| VerifiedSlideShapes {
                bytes: Arc::clone(&verified.bytes),
                persist_id: verified.persist_id,
                shapes: Arc::clone(&verified.shapes),
            })
    }

    /// Immutable source snapshot.
    #[must_use]
    pub(crate) const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Current number of slides.
    #[must_use]
    pub(crate) fn slide_count(&self) -> usize {
        self.document
            .slides()
            .map_or(self.source.slide_count(), |slides| slides.len())
    }

    /// Semantic moves staged in call order.
    #[must_use]
    pub(crate) fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Fixed-width slide visibility changes staged in call order.
    #[must_use]
    pub(crate) fn slide_visibility_changes(&self) -> &[SlideVisibilityChange] {
        &self.visibility_changes
    }

    /// Fixed-width slide advance changes staged in call order.
    #[must_use]
    pub(crate) fn slide_advance_changes(&self) -> &[SlideAdvanceChange] {
        &self.advance_changes
    }

    /// Fixed-width visual transition changes staged in call order.
    #[must_use]
    pub(crate) fn slide_transition_visual_changes(&self) -> &[SlideTransitionVisualChange] {
        &self.transition_visual_changes
    }

    /// Checked shape-text replacements staged in this root transaction.
    #[must_use]
    pub(crate) fn shape_text_changes(&self) -> &[ShapeTextChange] {
        &self.text_changes
    }

    /// Checked shape-anchor formatting replacements staged in this root.
    #[must_use]
    pub(crate) fn shape_anchor_changes(&self) -> &[ShapeAnchorChange] {
        &self.anchor_changes
    }

    /// External movie/audio path replacements staged in this root.
    #[must_use]
    pub(crate) fn external_media_path_changes(&self) -> &[ExternalMediaPathChange] {
        &self.media_path_changes
    }

    /// External-media playback replacements staged in this root.
    #[must_use]
    pub(crate) fn external_media_playback_changes(&self) -> &[ExternalMediaPlaybackChange] {
        &self.media_playback_changes
    }

    /// Moves a slide to a final zero-based position in the current projected
    /// order. Moving a slide to its current position is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when either position is absent, or a package
    /// error when the structural candidate is invalid.
    pub(crate) fn move_slide(&mut self, from: Position, destination: Position) -> Result<()> {
        self.require_position(from)?;
        self.require_position(destination)?;
        if from == destination {
            return Ok(());
        }
        self.require_change_allowed()?;
        let before_order = order_digest(&self.document)?;
        self.document
            .move_slide(from.get(), destination.get())
            .map_err(Error::from)?;
        let after_order = order_digest(&self.document)?;
        let change = Change {
            from,
            destination,
            before_order,
            after_order,
        };
        self.changes.push(change);
        self.structural.push(StructuralChange::Move(change));
        Ok(())
    }

    /// Reorders the main and title masters (record 0357a): the master list
    /// becomes `order`, each entry a current zero-based master position.
    ///
    /// Slides, notes and title masters reference masters by `masterId`, which
    /// is unchanged, so every reference keeps its master. The identity order
    /// is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidElementValue`] unless `order` is a permutation of
    /// the master positions; [`Refusal::Signed`] for a change on a signed
    /// source.
    pub(crate) fn reorder_masters(&mut self, order: &[usize]) -> Result<()> {
        let count = self.document.masters()?.len();
        let mut seen = vec![false; count];
        if order.len() != count
            || !order
                .iter()
                .all(|index| *index < count && !std::mem::replace(&mut seen[*index], true))
        {
            return Err(Error::Refused(Refusal::InvalidElementValue {
                element: "master order",
                reason: "it is not a permutation of the master positions",
            }));
        }
        if order
            .iter()
            .enumerate()
            .all(|(position, index)| position == *index)
        {
            return Ok(());
        }
        self.require_change_allowed()?;
        self.document.reorder_masters(order)?;
        self.push_document(DocumentChange::Masters {
            order: order.to_vec(),
        })?;
        Ok(())
    }

    /// Moves one master to a final zero-based position (record 0357a).
    ///
    /// # Errors
    ///
    /// [`Refusal::ElementNotFound`] for a position outside the master list.
    pub(crate) fn move_master(&mut self, from: Position, destination: Position) -> Result<()> {
        let count = self.document.masters()?.len();
        if from.get() >= count || destination.get() >= count {
            return Err(Error::Refused(Refusal::ElementNotFound {
                element: "master",
            }));
        }
        let mut order: Vec<_> = (0..count).collect();
        let value = order.remove(from.get());
        order.insert(destination.get(), value);
        self.reorder_masters(&order)
    }

    /// Moves the master with `masterId` `master_id` to a final position
    /// (record 0357a).
    ///
    /// # Errors
    ///
    /// [`Refusal::ElementNotFound`] when no master has that id.
    pub(crate) fn move_master_id(&mut self, master_id: u32, destination: Position) -> Result<()> {
        let from = self
            .document
            .masters()?
            .iter()
            .position(|master| master.master_id() == master_id)
            .ok_or(Error::Refused(Refusal::ElementNotFound {
                element: "master",
            }))?;
        self.move_master(Position::new(from), destination)
    }

    fn custom_table_styles(
        &self,
    ) -> Result<Option<document_structure::CustomTableStylesPlacement>> {
        Ok(self.document.structure()?.custom_table_styles)
    }

    fn require_custom_table_styles(
        &self,
    ) -> Result<document_structure::CustomTableStylesPlacement> {
        self.custom_table_styles()?
            .ok_or(Error::Refused(Refusal::ElementNotFound {
                element: "custom table styles",
            }))
    }

    /// Inserts a `RoundTripCustomTableStyles12Atom` before or after the
    /// `EndDocumentAtom` (record 0357a).
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidElementValue`] for a record that is not such an
    /// atom, and [`Refusal::UnsupportedElement`] when the document already
    /// has one.
    pub(crate) fn insert_custom_table_styles(
        &mut self,
        record: crate::Record,
        placement: document_structure::CustomTableStylesPlacement,
    ) -> Result<()> {
        require_custom_table_styles_record(&record)?;
        if self.custom_table_styles()?.is_some() {
            return Err(Error::Refused(Refusal::UnsupportedElement {
                element: "custom table styles",
                reason: "the document already has a custom table-style atom",
            }));
        }
        self.require_change_allowed()?;
        self.document
            .insert_custom_table_styles(record.clone(), placement)?;
        self.push_document(DocumentChange::InsertStyles { record, placement })?;
        Ok(())
    }

    /// Removes the custom table-style atom and returns it (record 0357a).
    ///
    /// # Errors
    ///
    /// [`Refusal::ElementNotFound`] when the document has none.
    pub(crate) fn remove_custom_table_styles(&mut self) -> Result<crate::Record> {
        let placement = self.require_custom_table_styles()?;
        self.require_change_allowed()?;
        let record = self.document.remove_custom_table_styles()?;
        self.push_document(DocumentChange::RemoveStyles {
            record: record.clone(),
            placement,
        })?;
        Ok(record)
    }

    /// Replaces the custom table-style atom in place and returns the
    /// previous one (record 0357a). An equal record is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::ElementNotFound`] when the document has none and
    /// [`Refusal::InvalidElementValue`] for a record that is not such an
    /// atom.
    pub(crate) fn replace_custom_table_styles(
        &mut self,
        record: crate::Record,
    ) -> Result<crate::Record> {
        require_custom_table_styles_record(&record)?;
        self.require_custom_table_styles()?;
        let mut probe = self.document.clone();
        let before = probe.replace_custom_table_styles(record.clone())?;
        if before == record {
            return Ok(before);
        }
        self.require_change_allowed()?;
        self.document = probe;
        self.push_document(DocumentChange::ReplaceStyles {
            before: before.clone(),
            after: record,
        })?;
        Ok(before)
    }

    /// Moves the custom table-style atom before or after the
    /// `EndDocumentAtom` (record 0357a). Its current placement is an exact
    /// no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::ElementNotFound`] when the document has none.
    pub(crate) fn move_custom_table_styles(
        &mut self,
        placement: document_structure::CustomTableStylesPlacement,
    ) -> Result<()> {
        let before = self.require_custom_table_styles()?;
        if before == placement {
            return Ok(());
        }
        self.require_change_allowed()?;
        self.document.move_custom_table_styles(placement)?;
        self.push_document(DocumentChange::MoveStyles {
            before,
            after: placement,
        })?;
        Ok(())
    }

    /// Sets one slide's `SlidePersistAtom.fNonOutlineData` flag
    /// ([MS-PPT] 2.4.14.5; record 0357a). It does not hide the slide. The
    /// current value is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::SlideNotFound`] for a position outside the slide list.
    pub(crate) fn set_slide_non_outline_data(
        &mut self,
        position: Position,
        present: bool,
    ) -> Result<()> {
        self.require_position(position)?;
        let slide = self.document.slides()?[position.get()];
        let before = slide.flags() & (1 << 2) != 0;
        if before == present {
            return Ok(());
        }
        self.require_change_allowed()?;
        self.document
            .set_slide_non_outline_data(position.get(), present)?;
        self.push_document(DocumentChange::NonOutlineData {
            position: position.get(),
            before,
            after: present,
        })?;
        Ok(())
    }

    /// Changes whether one slide is hidden during presentation.
    ///
    /// Only `SlideShowSlideInfoAtom.fHidden` (flags bit 2), defined by
    /// `[MS-PPT]` 2.6.6, is changed. The similarly positioned
    /// `SlidePersistAtom.fNonOutlineData` bit is not a visibility flag and is
    /// preserved exactly. The operation composes with order, shape, and inert
    /// media changes and is carried by exact, inverse, durable, history, and
    /// merge surfaces.
    ///
    /// A slide without a `SlideShowSlideInfoAtom` (producer slides, and
    /// slides the PPT writer made without timing or a transition) gains one
    /// directly after its `SlideAtom` ([MS-PPT] 2.5.1) when it is hidden: no
    /// transition, no sound, manual advance, medium speed, `fHidden` set.
    /// Showing such a slide is a no-op. Clearing the flag again later keeps
    /// the atom; the exact inverse of the commit restores the source bytes.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the position is absent, or the slide has
    /// several or a malformed slide-show information atom, or no leading
    /// `SlideAtom` to insert one after.
    pub(crate) fn set_slide_hidden(&mut self, position: Position, hidden: bool) -> Result<()> {
        self.require_position(position)?;
        let selected = self.document.slides()?[position.get()];
        let source_position = self
            .source
            .document
            .slides()
            .iter()
            .position(|slide| slide.persist_id() == selected.persist_id())
            .map(Position::new)
            .ok_or(Error::Refused(Refusal::UncommittedSlideDependency))?;
        let before = self.working.slide_hidden(source_position)?;
        if before == hidden {
            return Ok(());
        }
        self.require_change_allowed()?;
        self.working = replace_slide_hidden(&self.working, source_position, hidden)?;
        let change = SlideVisibilityChange {
            position,
            target: source_position,
            slide_id: selected.slide_id(),
            before,
            after: hidden,
        };
        self.visibility_changes.push(change);
        self.formatting
            .push(FormattingChange::SlideVisibility(change));
        Ok(())
    }

    /// Replaces one existing slide's manual/automatic advance state.
    ///
    /// This changes only `slideTime`, `fManualAdvance`, and `fAutoAdvance` in
    /// the fixed-width `SlideShowSlideInfoAtom`. Transition visuals, sounds,
    /// hidden/cursor flags, reserved bits, and every other persisted record
    /// remain exact.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an absent, ambiguous, malformed, or newly
    /// inserted slide-show information owner.
    pub(crate) fn set_slide_advance(
        &mut self,
        position: Position,
        advance: SlideAdvance,
    ) -> Result<()> {
        self.require_position(position)?;
        let selected = self.document.slides()?[position.get()];
        let source_position = self
            .source
            .document
            .slides()
            .iter()
            .position(|slide| slide.persist_id() == selected.persist_id())
            .map(Position::new)
            .ok_or(Error::Refused(Refusal::UncommittedSlideDependency))?;
        let before = self.working.slide_advance(source_position)?;
        if before == advance {
            return Ok(());
        }
        self.require_change_allowed()?;
        self.working = replace_slide_advance(&self.working, source_position, advance)?;
        let change = SlideAdvanceChange {
            target: source_position,
            slide_id: selected.slide_id(),
            before,
            after: advance,
        };
        self.advance_changes.push(change);
        self.formatting.push(FormattingChange::SlideAdvance(change));
        Ok(())
    }

    /// Replaces one existing slide's visual transition kind, direction, and speed.
    ///
    /// Timing, sound reference/state, all transition flags, unused bytes, and
    /// every other persisted record remain exact.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an absent, ambiguous, malformed, or newly
    /// inserted visual transition owner.
    pub(crate) fn set_slide_transition_visual(
        &mut self,
        position: Position,
        visual: SlideTransitionVisual,
    ) -> Result<()> {
        self.require_position(position)?;
        let selected = self.document.slides()?[position.get()];
        let source_position = self
            .source
            .document
            .slides()
            .iter()
            .position(|slide| slide.persist_id() == selected.persist_id())
            .map(Position::new)
            .ok_or(Error::Refused(Refusal::UncommittedSlideDependency))?;
        let before = self.working.slide_transition_visual(source_position)?;
        if before == visual {
            return Ok(());
        }
        self.require_change_allowed()?;
        self.working = replace_slide_transition_visual(&self.working, source_position, visual)?;
        let change = SlideTransitionVisualChange {
            target: source_position,
            slide_id: selected.slide_id(),
            before,
            after: visual,
        };
        self.transition_visual_changes.push(change);
        self.formatting
            .push(FormattingChange::SlideTransitionVisual(change));
        Ok(())
    }

    /// Replaces text in an existing shape while preserving its formatting and
    /// validating the shape's modeled dependency closure.
    ///
    /// The target slide position is resolved against the transaction's
    /// projected structural order. The text publication remains isolated
    /// until this root transaction is committed.
    ///
    /// # Errors
    ///
    /// Returns a typed slide refusal or the checked shape-text owner's error.
    pub(crate) fn set_shape_text(
        &mut self,
        target: crate::text_edit::Target,
        value: impl Into<String>,
    ) -> Result<()> {
        self.require_position(target.slide())?;
        let projected = self.document.slides()?;
        let persist_id = projected[target.slide().get()].persist_id();
        let source_position = self
            .source
            .document
            .slides()
            .iter()
            .position(|slide| slide.persist_id() == persist_id)
            .map(Position::new)
            .ok_or(Error::Refused(Refusal::UncommittedSlideDependency))?;
        let source_target = crate::text_edit::Target::new(source_position, target.shape());
        let after = value.into();
        if self.source.has_signature {
            let before = crate::text_edit::inspect_shape_text_readonly(
                self.working.bytes(),
                source_target,
                self.source.open_options(),
            )?;
            if before == after {
                return Ok(());
            }
            return Err(Error::Refused(Refusal::Signed));
        }
        let text_snapshot =
            crate::text_edit::Snapshot::from_shared_bytes(self.working.shared_bytes())?;
        let mut text_edit = match text_snapshot.edit_text_in(source_target, self.opened_source()) {
            Ok(text_edit) => text_edit,
            Err(crate::text_edit::Error::Refused(crate::text_edit::Refusal::NoTextAtom)) => {
                // Outline-backed placeholder text is owned by the live
                // document; the batch owner publishes that body.
                return self
                    .set_shape_texts_mapped(&[(target, after.as_str())])
                    .map(drop);
            },
            Err(error) => return Err(error.into()),
        };
        let before = text_edit.text().to_string();
        text_edit.set_text(after.clone())?;
        let (commit, verified) = text_edit.commit_verified()?;
        if commit.patch().is_empty() {
            return Ok(());
        }
        let publication = commit.into_root_publication().ok_or_else(|| {
            PackageError::Corrupted("changed PPT text commit lacks publication identity".into())
        })?;
        self.working = self
            .working
            .adopt_text_edit_publication(publication, persist_id)?;
        // As for batches: proven shapes describe exactly the adopted bytes.
        self.verified_shapes = verified.map(|(persist_id, shapes)| VerifiedShapes {
            bytes: self.working.shared_bytes(),
            persist_id,
            shapes: Arc::new(shapes),
        });
        let change = ShapeTextChange {
            target: source_target,
            before,
            after,
        };
        self.text_changes.push(change.clone());
        self.formatting.push(FormattingChange::Text(change));
        Ok(())
    }

    /// Replaces the speaker-notes text of the slide at a projected position.
    ///
    /// The slide's notes page must hold exactly one body placeholder; its
    /// text body follows the shape-text dependency-closure rules. Equal text
    /// is an exact no-op. The change publishes the notes persist object in
    /// one append-only user edit and reopens the candidate for readback.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for a missing slide or notes page, a notes
    /// page without one body placeholder, an unrepresentable or
    /// position-dependent text change, a changed signed source, or a slide
    /// inserted by this transaction.
    pub(crate) fn set_notes_text(
        &mut self,
        position: Position,
        value: impl Into<String>,
    ) -> Result<()> {
        self.require_position(position)?;
        let projected = self.document.slides()?;
        let persist_id = projected[position.get()].persist_id();
        let source_position = self
            .source
            .document
            .slides()
            .iter()
            .position(|slide| slide.persist_id() == persist_id)
            .map(Position::new)
            .ok_or(Error::Refused(Refusal::UncommittedSlideDependency))?;
        let after = value.into();
        if self.source.has_signature {
            let before = crate::text_edit::inspect_notes_text_readonly(
                self.working.bytes(),
                source_position,
                self.source.open_options(),
            )?;
            if before == after {
                return Ok(());
            }
            return Err(Error::Refused(Refusal::Signed));
        }
        let Some(publication) = crate::text_edit::replace_notes_text(
            self.working.shared_bytes(),
            source_position,
            &after,
            self.source.limits.max_package_bytes,
        )?
        else {
            return Ok(());
        };
        let before = publication.before().to_owned();
        self.working = self.working.adopt_notes_publication(publication)?;
        let change = NotesTextChange {
            slide: source_position,
            before,
            after,
        };
        self.formatting.push(FormattingChange::NotesText(change));
        Ok(())
    }

    /// Atomically replaces text in several existing native PPT shapes.
    ///
    /// All projected targets are resolved against the immutable working root
    /// captured at entry. Encodings and dependency closures are validated for
    /// the complete batch before publication. Changed shapes are grouped by
    /// persisted slide, and all changed slide records are emitted through one
    /// append-only PPT user edit. Semantic changes retain caller order; exact
    /// no-ops are omitted and preserve the working source allocation.
    ///
    /// # Errors
    ///
    /// Returns a resource-limit error before proportional work when more than
    /// [`MAX_SHAPE_TEXT_REPLACEMENTS`] operations are supplied. Duplicate
    /// semantic targets, missing targets, incompatible encodings, unsafe
    /// dependency closures, and package publication failures leave this
    /// transaction unchanged.
    pub(crate) fn set_shape_texts(
        &mut self,
        replacements: &[ShapeTextReplacement],
    ) -> Result<usize> {
        if replacements.len() > MAX_SHAPE_TEXT_REPLACEMENTS {
            return Err(PackageError::ResourceLimit(format!(
                "PPT shape-text batch operation count {} exceeds limit {}",
                replacements.len(),
                MAX_SHAPE_TEXT_REPLACEMENTS
            ))
            .into());
        }
        if replacements.is_empty() {
            return Ok(0);
        }

        let mapped = replacements
            .iter()
            .map(|replacement| (replacement.target(), replacement.text()))
            .collect::<Vec<_>>();
        self.set_shape_texts_mapped(&mapped)
    }

    fn set_shape_texts_mapped(
        &mut self,
        replacements: &[(crate::text_edit::Target, &str)],
    ) -> Result<usize> {
        if replacements.is_empty() {
            return Ok(0);
        }

        let projected = self.document.slides()?;
        let mut seen = BTreeSet::new();
        let mut mapped = Vec::with_capacity(replacements.len());
        let mut expected_persist_ids = BTreeMap::new();
        for (target, replacement) in replacements {
            if !seen.insert((target.slide().get(), target.shape().get())) {
                return Err(Error::Refused(Refusal::DuplicateShapeTextTarget {
                    target: *target,
                }));
            }
            let selected = projected.get(target.slide().get()).ok_or(Error::Refused(
                Refusal::SlideNotFound {
                    position: target.slide(),
                },
            ))?;
            let persist_id = selected.persist_id();
            let source_position = self
                .source
                .document
                .slides()
                .iter()
                .position(|slide| slide.persist_id() == persist_id)
                .map(Position::new)
                .ok_or(Error::Refused(Refusal::UncommittedSlideDependency))?;
            let source_target = crate::text_edit::Target::new(source_position, target.shape());
            mapped.push((source_target, *replacement));
            expected_persist_ids.insert(
                (source_target.slide().get(), source_target.shape().get()),
                persist_id,
            );
        }

        if self.source.has_signature {
            for (target, replacement) in &mapped {
                let before = crate::text_edit::inspect_shape_text_readonly(
                    self.working.bytes(),
                    *target,
                    self.source.open_options(),
                )?;
                if before != *replacement {
                    return Err(Error::Refused(Refusal::Signed));
                }
            }
            return Ok(0);
        }

        let Some(mut publication) = crate::text_edit::replace_shape_texts_batch(
            self.working.shared_bytes(),
            &mapped,
            self.source.limits.max_package_bytes,
            self.opened_source(),
        )?
        else {
            return Ok(0);
        };

        // Validate the native publication identities against the projected
        // targets before adopting either bytes or semantic patch state.
        for change in publication.changes() {
            let key = (change.target().slide().get(), change.target().shape().get());
            if expected_persist_ids.get(&key).copied() != Some(change.slide_persist_id()) {
                return Err(PackageError::Corrupted(
                    "PPT batch text publication targeted an unexpected slide record".into(),
                )
                .into());
            }
        }

        let document_changed = publication.document_changed();
        let verified = publication.take_verified_shapes();
        let (working, published_changes) = self
            .working
            .adopt_text_edit_batch_publication(publication)?;
        // Proven shapes describe exactly the adopted bytes when adoption kept
        // the publication's allocation; a reopened adoption drops them.
        self.verified_shapes = verified
            .into_iter()
            .next_back()
            .map(|(persist_id, shapes)| VerifiedShapes {
                bytes: working.shared_bytes(),
                persist_id,
                shapes: Arc::new(shapes),
            });
        self.live_document_changed |= document_changed;
        let changes = published_changes
            .into_iter()
            .map(|change| ShapeTextChange {
                target: change.target(),
                before: change.before().to_string(),
                after: change.after().to_string(),
            })
            .collect::<Vec<_>>();
        let changed = changes.len();

        self.working = working;
        self.text_changes.extend(changes.iter().cloned());
        self.formatting
            .extend(changes.into_iter().map(FormattingChange::Text));
        Ok(changed)
    }

    /// Atomically replaces shape text from several bounded replacement
    /// chunks.
    ///
    /// Each inner chunk retains the 256-target contract of
    /// [`Self::set_shape_texts`]. The aggregate accepts at most
    /// [`MAX_SHAPE_TEXT_CHUNKS`] chunks, including empty chunks, and at most
    /// [`MAX_SHAPE_TEXT_CHUNK_SELECTIONS`] selected shapes. Targets must be
    /// globally distinct across all chunks. The complete flattened plan is
    /// admitted before target resolution, then all changed slide records are
    /// grouped into one native publication and one candidate reopen/readback.
    /// Exact no-ops retain the working allocation, and a failed call leaves
    /// this transaction unchanged.
    ///
    /// # Errors
    ///
    /// Returns a resource-limit error for a chunk, aggregate selection, or
    /// replacement-byte limit; duplicate semantic targets, missing targets,
    /// incompatible encodings, unsafe dependency closures, signed changes,
    /// and publication failures are returned without partial state.
    pub(crate) fn set_shape_text_chunks(
        &mut self,
        chunks: &[&[ShapeTextReplacement]],
    ) -> Result<usize> {
        let (flattened, replacement_bytes) = flatten_shape_text_chunks(chunks)?;
        if replacement_bytes > self.source.limits.max_package_bytes {
            return Err(PackageError::ResourceLimit(format!(
                "PPT shape-text aggregate replacement bytes {} exceeds limit {}",
                replacement_bytes, self.source.limits.max_package_bytes
            ))
            .into());
        }
        if flattened.is_empty() {
            return Ok(0);
        }
        self.set_shape_texts_mapped(&flattened)
    }

    /// Replaces an existing shape's complete checked `PowerPoint` client anchor.
    /// Every other `OfficeArt` record and CFB stream remains exact.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the semantic target is missing,
    /// ambiguous, newly inserted, or lacks exactly one supported host anchor.
    pub(crate) fn set_shape_anchor(
        &mut self,
        target: crate::text_edit::Target,
        anchor: crate::Anchor,
    ) -> Result<()> {
        self.require_position(target.slide())?;
        let projected = self.document.slides()?;
        let persist_id = projected[target.slide().get()].persist_id();
        let source_position = self
            .source
            .document
            .slides()
            .iter()
            .position(|slide| slide.persist_id() == persist_id)
            .map(Position::new)
            .ok_or(Error::Refused(Refusal::UncommittedSlideDependency))?;
        let source_target = crate::text_edit::Target::new(source_position, target.shape());
        if self.source.has_signature {
            let before = crate::text_edit::inspect_shape_anchor_readonly(
                self.working.bytes(),
                source_target,
                self.source.open_options(),
            )?;
            if before == anchor {
                return Ok(());
            }
            return Err(Error::Refused(Refusal::Signed));
        }
        let rewritten =
            crate::text_edit::replace_shape_anchor(self.working.bytes(), source_target, anchor)?;
        if rewritten.before == rewritten.after {
            return Ok(());
        }
        self.working =
            Snapshot::from_bytes_with_options(rewritten.bytes, self.source.open_options())?;
        let change = ShapeAnchorChange {
            target: source_target,
            before: rewritten.before,
            after: rewritten.after,
        };
        self.anchor_changes.push(change);
        self.formatting.push(FormattingChange::Anchor(change));
        Ok(())
    }

    /// Changes or clears an inert linked movie/audio path through the live
    /// document's external-media owner.
    ///
    /// The path is serialized only; it is never opened or resolved. A
    /// transaction that also stages slide-list structure is rebased over the
    /// changed live document before its single append-only publication.
    ///
    /// # Errors
    ///
    /// Returns a typed owner-order refusal or an external-media validation or
    /// package publication error.
    pub(crate) fn set_external_media_path(&mut self, id: u32, path: Option<String>) -> Result<()> {
        let media = external_media_snapshot(&self.working)?;
        let object = media
            .collection()
            .and_then(|collection| collection.objects.iter().find(|object| object.id() == id))
            .ok_or_else(|| {
                PackageError::InvalidFormat(format!("external media ID {id} was not found"))
            })?;
        let before = media_path_of(object)?;
        if self.source.has_signature {
            if before.as_deref() == path.as_deref() {
                return Ok(());
            }
            return Err(Error::Refused(Refusal::Signed));
        }
        let mut transaction = media.edit();
        transaction.set_path(id, path.clone())?;
        let commit = transaction.commit()?;
        if commit.patch().is_empty() {
            return Ok(());
        }
        self.working = publish_live_document(&self.working, commit.snapshot().bytes())?;
        if self.working.external_media_path(id)? != path {
            return Err(PackageError::Corrupted(
                "published PPT external-media path did not round-trip".into(),
            )
            .into());
        }
        let change = ExternalMediaPathChange {
            id,
            before,
            after: path,
        };
        self.media_path_changes.push(change.clone());
        self.formatting.push(FormattingChange::MediaPath(change));
        Ok(())
    }

    /// Replaces playback flags through the live document's external-media
    /// owner without resolving or activating external content.
    ///
    /// # Errors
    ///
    /// Returns a typed owner-order refusal or an external-media validation or
    /// package publication error.
    pub(crate) fn set_external_media_playback(
        &mut self,
        id: u32,
        playback: crate::external_media::Playback,
    ) -> Result<()> {
        let media = external_media_snapshot(&self.working)?;
        let before = media
            .collection()
            .and_then(|collection| collection.playback(id))
            .ok_or_else(|| {
                PackageError::InvalidFormat(format!("external media ID {id} was not found"))
            })?;
        if self.source.has_signature {
            if before == playback {
                return Ok(());
            }
            return Err(Error::Refused(Refusal::Signed));
        }
        let mut transaction = media.edit();
        transaction.set_playback(id, playback)?;
        let commit = transaction.commit()?;
        if commit.patch().is_empty() {
            return Ok(());
        }
        self.working = publish_live_document(&self.working, commit.snapshot().bytes())?;
        if self.working.external_media_playback(id)? != playback {
            return Err(PackageError::Corrupted(
                "published PPT external-media playback did not round-trip".into(),
            )
            .into());
        }
        let change = ExternalMediaPlaybackChange {
            id,
            before,
            after: playback,
        };
        self.media_playback_changes.push(change);
        self.formatting
            .push(FormattingChange::MediaPlayback(change));
        Ok(())
    }

    /// Removes one slide-list entry and its dependency-owned outline records.
    /// The persisted slide payload remains as unreachable incremental history.
    ///
    /// # Errors
    ///
    /// Returns a typed position refusal or a structural validation error.
    pub(crate) fn remove_slide(&mut self, position: Position) -> Result<()> {
        self.require_position(position)?;
        self.require_change_allowed()?;
        let before_order = order_digest(&self.document)?;
        let slides = self.document.slides()?;
        let selected = slides[position.get()];
        if self.inserted_records.contains_key(&selected.persist_id()) {
            return Err(Error::Refused(Refusal::UncommittedSlideDependency));
        }
        let group = self.document.remove_slide(position.get())?;
        let after_order = order_digest(&self.document)?;
        let record = persisted_record(&self.working, selected.persist_id())?;
        let payload = slide_payload(group, record)?;
        self.structural.push(StructuralChange::Remove(ListChange {
            position,
            before_order,
            after_order,
            payload,
        }));
        Ok(())
    }

    /// Inserts a blank slide at a projected position; `position ==
    /// slide_count()` appends.
    ///
    /// The slide uses the blank layout and follows the master, color scheme
    /// and background of its preceding slide (or the following slide when it
    /// becomes the first, or the first main master of an empty deck). Its
    /// drawing receives a new drawing identifier and shape-identifier
    /// cluster, registered in the drawing group at commit.
    ///
    /// # Errors
    ///
    /// Returns a typed position or signed-source refusal, or a resource
    /// error when the persist, slide, drawing or shape identifier space is
    /// exhausted.
    pub(crate) fn insert_slide(&mut self, position: Position) -> Result<()> {
        self.require_change_allowed()?;
        let count = self.slide_count();
        if position.get() > count {
            return Err(Error::Refused(Refusal::SlideNotFound { position }));
        }
        let slides = self.document.slides()?;
        let neighbor = position
            .get()
            .checked_sub(1)
            .or_else(|| (count > 0).then_some(0))
            .map(|index| slides[index].persist_id());
        let (master_id, color_scheme) = match neighbor {
            Some(persist_id) => {
                let record = match self.inserted_records.get(&persist_id) {
                    Some(record) => record.clone(),
                    None => persisted_record(&self.working, persist_id)?,
                };
                blank::master_and_scheme(&record)?
            },
            None => {
                let master = self
                    .document
                    .masters()?
                    .first()
                    .copied()
                    .ok_or(Error::Refused(Refusal::MissingMasterDependency {
                        master_id: 0,
                    }))?;
                (master.master_id(), None)
            },
        };
        require_master(&self.document, master_id)?;
        let allocation = blank::allocate(self.working.document.bytes(), &self.inserted_records)?;
        let record = blank::slide_record(master_id, color_scheme, allocation)?;
        let payload = SlidePayload {
            group: blank::slide_list_group(),
            record,
            master_id,
            pictures: Vec::new(),
        };
        self.insert_payload(position, payload)
    }

    /// Inserts a dependency-checked cross-presentation transfer plan.
    /// `position == slide_count()` appends the transferred slide.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for a stale/different target or a package error
    /// when identifier allocation or structural validation fails.
    pub(crate) fn insert_transfer(
        &mut self,
        position: Position,
        plan: &TransferPlan,
    ) -> Result<()> {
        self.require_change_allowed()?;
        if plan.target_artifact != self.source.artifact_digest() {
            return Err(Error::Refused(Refusal::TransferTargetMismatch));
        }
        if position.get() > self.slide_count() {
            return Err(Error::Refused(Refusal::SlideNotFound { position }));
        }
        require_master(&self.document, plan.payload.master_id)?;
        let picture_slots = picture_inventory(&self.source)?.slots;
        for (index, picture) in plan.payload.pictures.iter().enumerate() {
            let expected_id = picture_slots
                .checked_add(u32::try_from(index).map_err(|_error| {
                    PackageError::ResourceLimit("too many transferred PPT pictures".into())
                })?)
                .and_then(|value| value.checked_add(1))
                .ok_or_else(|| PackageError::ResourceLimit("PPT BLIP ID overflow".into()))?;
            if picture.target_id != expected_id {
                return Err(PackageError::InvalidFormat(
                    "PPT transfer picture IDs are not contiguous after the target BStore".into(),
                )
                .into());
            }
            build_picture_fbse(&picture.blip, 0)?;
            match self.inserted_pictures.get(&picture.target_id) {
                Some(existing) if existing == &picture.blip => {},
                Some(_) => {
                    return Err(Error::Refused(Refusal::UncommittedSlideDependency));
                },
                None => {
                    self.inserted_pictures
                        .insert(picture.target_id, picture.blip.clone());
                },
            }
        }
        self.insert_payload(position, plan.payload.clone())
    }

    /// Allocates the persist and slide identifiers of an inserted slide and
    /// stages its slide-list group and persisted record.
    fn insert_payload(&mut self, position: Position, payload: SlidePayload) -> Result<()> {
        let persist_id = next_persist_id(&self.source, &self.inserted_records)?;
        let slide_id = next_slide_id(&self.document)?;
        let mut group = payload.group;
        rewrite_slide_persist(&mut group, persist_id, slide_id)?;
        let before_order = order_digest(&self.document)?;
        self.document
            .insert_slide_group(position.get(), group.clone())?;
        let after_order = order_digest(&self.document)?;
        self.inserted_records
            .insert(persist_id, payload.record.clone());
        self.structural.push(StructuralChange::Insert(ListChange {
            position,
            before_order,
            after_order,
            payload: SlidePayload {
                group,
                record: payload.record,
                master_id: payload.master_id,
                pictures: payload.pictures,
            },
        }));
        Ok(())
    }

    /// Publishes the complete order atomically through an append-only PPT
    /// user edit and reopens the whole candidate before returning it.
    ///
    /// # Errors
    ///
    /// Returns an error if publication, preservation checks, full reopen, or
    /// semantic readback fails.
    pub(crate) fn commit(self) -> Result<Commit> {
        if self.source.has_signature && self.source.bytes() != self.working.bytes() {
            return Err(Error::Refused(Refusal::Signed));
        }
        let document_commit = if self.structural.is_empty()
            || (self.media_path_changes.is_empty()
                && self.media_playback_changes.is_empty()
                && !self.live_document_changed)
        {
            self.document.commit()?
        } else {
            let mut rebased = self.working.document.edit();
            replay_structural_changes(&mut rebased, &self.structural)?;
            rebased.commit()?
        };
        let source = self.source;
        let working = self.working;
        if source.has_signature
            && (!document_commit.patch().is_empty() || source.bytes() != working.bytes())
        {
            return Err(Error::Refused(Refusal::Signed));
        }
        if document_commit.patch().is_empty() {
            // Any staged structural operations net to the published working
            // artifact, which the patch retains as its target.
            let patch = Patch {
                before: source,
                after: working.clone(),
                changes: self.changes,
                structural: self.structural,
                visibility_changes: self.visibility_changes,
                advance_changes: self.advance_changes,
                transition_visual_changes: self.transition_visual_changes,
                text_changes: self.text_changes,
                anchor_changes: self.anchor_changes,
                media_path_changes: self.media_path_changes,
                media_playback_changes: self.media_playback_changes,
                formatting: self.formatting,
                formatting_first: true,
                structural_before: StructuralArtifact::After,
                structural_after: StructuralArtifact::After,
            };
            return Ok(Commit {
                snapshot: working,
                patch,
            });
        }

        let mut editor = crate::embedded::object::Editor::open_records_arc_with_limit(
            working.shared_bytes(),
            source.limits.max_package_bytes,
            source.limits,
        )?;
        // Nothing is staged yet, so the publishing editor reads exactly the
        // working artifact's persisted slide payloads.
        let before_slides = editor_persisted_slides(&editor, &working.document)?;
        let live = editor.persisted_record(source.document_persist_id)?;
        if live.as_slice() != working.document.bytes() {
            return Err(PackageError::Corrupted(
                "PPT slide-order transaction source changed before publication".into(),
            )
            .into());
        }
        for (persist_id, record) in &self.inserted_records {
            editor.insert_persisted_record(*persist_id, record.clone())?;
        }
        let mut document_bytes = blank::register_inserted_drawings(
            document_commit.snapshot().bytes(),
            &self.inserted_records,
        )?;
        let picture_stream = if self.inserted_pictures.is_empty() {
            None
        } else {
            let (rewritten, pictures) =
                append_transferred_pictures(&working, &document_bytes, &self.inserted_pictures)?;
            document_bytes = rewritten;
            Some(pictures)
        };
        editor.replace_persisted_record(source.document_persist_id, document_bytes)?;
        if let Some(pictures) = picture_stream {
            editor.replace_stream_by_name("Pictures", pictures)?;
        }
        let bytes = editor.finish()?;
        if self.inserted_pictures.is_empty() {
            crate::font::validate_unrelated_streams(working.bytes(), &bytes)?;
        } else {
            validate_unrelated_picture_publication(working.bytes(), &bytes)?;
        }
        let snapshot = Snapshot::from_bytes_with_options(bytes, source.open_options())?;
        if snapshot.document.slides() != document_commit.snapshot().slides() {
            return Err(PackageError::Corrupted(
                "published PPT slide order did not round-trip through the live document".into(),
            )
            .into());
        }
        let after_slides = persisted_slides(
            snapshot.artifact.shared(),
            &snapshot.document,
            snapshot.limits,
        )?;
        let expected_slides = expected_persisted_slides(
            document_commit.snapshot().slides(),
            &before_slides,
            &self.inserted_records,
        )?;
        if after_slides != expected_slides {
            return Err(PackageError::Corrupted(
                "PPT slide-order publication changed a slide payload record".into(),
            )
            .into());
        }
        validate_transferred_pictures(&snapshot, &self.inserted_pictures)?;
        let structural_before = StructuralArtifact::structural_source(&source, &working);
        let patch = Patch {
            before: source,
            after: snapshot.clone(),
            changes: self.changes,
            structural: self.structural,
            visibility_changes: self.visibility_changes,
            advance_changes: self.advance_changes,
            transition_visual_changes: self.transition_visual_changes,
            text_changes: self.text_changes,
            anchor_changes: self.anchor_changes,
            media_path_changes: self.media_path_changes,
            media_playback_changes: self.media_playback_changes,
            formatting: self.formatting,
            formatting_first: true,
            structural_before,
            structural_after: StructuralArtifact::After,
        };
        Ok(Commit { snapshot, patch })
    }

    /// Publishes the transaction while reporting deterministic, content-free
    /// phase boundaries to an external observer.
    ///
    /// This is an opt-in diagnostic equivalent of [`Self::commit`]. It keeps
    /// the ordinary commit implementation and all of its validation and
    /// ownership rules unchanged. The observer is called synchronously in
    /// operation order; this method does not timestamp events, retain event
    /// data, or consult ambient runtime state. Each started phase is closed by
    /// one matching finished event, including a phase that returns an error.
    ///
    /// # Errors
    ///
    /// Returns the same typed error as [`Self::commit`].
    #[cfg(feature = "performance-diagnostics")]
    pub(crate) fn commit_profiled(
        self,
        mut observer: impl FnMut(DiagnosticEvent),
    ) -> Result<Commit> {
        if self.source.has_signature && self.source.bytes() != self.working.bytes() {
            return Err(Error::Refused(Refusal::Signed));
        }
        let document_commit = if self.structural.is_empty()
            || (self.media_path_changes.is_empty()
                && self.media_playback_changes.is_empty()
                && !self.live_document_changed)
        {
            observe_phase(&mut observer, DiagnosticPhase::DocumentCommit, || {
                Ok(self.document.commit()?)
            })?
        } else {
            observe_phase(&mut observer, DiagnosticPhase::DocumentCommit, || {
                let mut rebased = self.working.document.edit();
                replay_structural_changes(&mut rebased, &self.structural)?;
                Ok(rebased.commit()?)
            })?
        };
        let source = self.source;
        let working = self.working;
        if source.has_signature
            && (!document_commit.patch().is_empty() || source.bytes() != working.bytes())
        {
            return Err(Error::Refused(Refusal::Signed));
        }

        if document_commit.patch().is_empty() {
            observe_phase(&mut observer, DiagnosticPhase::StructuralNoOp, || Ok(()))?;
            let patch = Patch {
                before: source,
                after: working.clone(),
                changes: self.changes,
                structural: self.structural,
                visibility_changes: self.visibility_changes,
                advance_changes: self.advance_changes,
                transition_visual_changes: self.transition_visual_changes,
                text_changes: self.text_changes,
                anchor_changes: self.anchor_changes,
                media_path_changes: self.media_path_changes,
                media_playback_changes: self.media_playback_changes,
                formatting: self.formatting,
                formatting_first: true,
                structural_before: StructuralArtifact::After,
                structural_after: StructuralArtifact::After,
            };
            return Ok(Commit {
                snapshot: working,
                patch,
            });
        }

        let mut editor = observe_phase(&mut observer, DiagnosticPhase::EmbeddedOpen, || {
            Ok(
                crate::embedded::object::Editor::open_records_arc_with_limit(
                    working.shared_bytes(),
                    source.limits.max_package_bytes,
                    source.limits,
                )?,
            )
        })?;
        let before_slides =
            observe_phase(&mut observer, DiagnosticPhase::BeforePayloadCapture, || {
                editor_persisted_slides(&editor, &working.document)
            })?;
        let _live = observe_phase(&mut observer, DiagnosticPhase::LiveDocumentRead, || {
            let live = editor.persisted_record(source.document_persist_id)?;
            if live.as_slice() != working.document.bytes() {
                return Err(PackageError::Corrupted(
                    "PPT slide-order transaction source changed before publication".into(),
                )
                .into());
            }
            Ok(live)
        })?;
        for (persist_id, record) in &self.inserted_records {
            editor.insert_persisted_record(*persist_id, record.clone())?;
        }
        let mut document_bytes = blank::register_inserted_drawings(
            document_commit.snapshot().bytes(),
            &self.inserted_records,
        )?;
        let picture_stream = if self.inserted_pictures.is_empty() {
            None
        } else {
            let (rewritten, pictures) =
                append_transferred_pictures(&working, &document_bytes, &self.inserted_pictures)?;
            document_bytes = rewritten;
            Some(pictures)
        };
        editor.replace_persisted_record(source.document_persist_id, document_bytes)?;
        if let Some(pictures) = picture_stream {
            editor.replace_stream_by_name("Pictures", pictures)?;
        }
        let bytes = observe_phase(&mut observer, DiagnosticPhase::EmbeddedFinish, || {
            Ok(editor.finish()?)
        })?;
        observe_phase(
            &mut observer,
            DiagnosticPhase::UnrelatedStreamValidation,
            || {
                if self.inserted_pictures.is_empty() {
                    crate::font::validate_unrelated_streams(working.bytes(), &bytes)?;
                } else {
                    validate_unrelated_picture_publication(working.bytes(), &bytes)?;
                }
                Ok(())
            },
        )?;
        let snapshot = observe_phase(&mut observer, DiagnosticPhase::PublicReopen, || {
            Snapshot::from_bytes_with_options(bytes, source.open_options())
        })?;
        if snapshot.document.slides() != document_commit.snapshot().slides() {
            return Err(PackageError::Corrupted(
                "published PPT slide order did not round-trip through the live document".into(),
            )
            .into());
        }
        let after_slides =
            observe_phase(&mut observer, DiagnosticPhase::AfterPayloadCapture, || {
                persisted_slides(
                    snapshot.artifact.shared(),
                    &snapshot.document,
                    snapshot.limits,
                )
            })?;
        let expected_slides = expected_persisted_slides(
            document_commit.snapshot().slides(),
            &before_slides,
            &self.inserted_records,
        )?;
        if after_slides != expected_slides {
            return Err(PackageError::Corrupted(
                "PPT slide-order publication changed a slide payload record".into(),
            )
            .into());
        }
        validate_transferred_pictures(&snapshot, &self.inserted_pictures)?;
        let structural_before = if StructuralArtifact::retains_structural_source(&source, &working)
        {
            StructuralArtifact::Before
        } else {
            observe_phase(
                &mut observer,
                DiagnosticPhase::IntermediateArtifactHash,
                || Ok(StructuralArtifact::intermediate(&working)),
            )?
        };
        let patch = Patch {
            before: source,
            after: snapshot.clone(),
            changes: self.changes,
            structural: self.structural,
            visibility_changes: self.visibility_changes,
            advance_changes: self.advance_changes,
            transition_visual_changes: self.transition_visual_changes,
            text_changes: self.text_changes,
            anchor_changes: self.anchor_changes,
            media_path_changes: self.media_path_changes,
            media_playback_changes: self.media_playback_changes,
            formatting: self.formatting,
            formatting_first: true,
            structural_before,
            structural_after: StructuralArtifact::After,
        };
        Ok(Commit { snapshot, patch })
    }

    fn require_position(&self, position: Position) -> Result<()> {
        if position.get() >= self.slide_count() {
            return Err(Error::Refused(Refusal::SlideNotFound { position }));
        }
        Ok(())
    }

    /// Stage one element edit over persisted records (record 0323).
    ///
    /// `slide` selects a slide by its projected position; `document` offers
    /// the live document. The closure computes the complete replacement of
    /// each record it changes from the exact current bytes. An empty
    /// replacement is an exact no-op; a signed source refuses a changed
    /// one. Returns whether anything changed.
    pub(crate) fn stage_records(
        &mut self,
        slide: Option<Position>,
        document: bool,
        edit: impl FnOnce(records::View<'_>) -> Result<records::Replacement>,
    ) -> Result<bool> {
        let slide = match slide {
            Some(position) => {
                self.require_position(position)?;
                let selected = self.document.slides()?[position.get()];
                let source_position = self
                    .source
                    .document
                    .slides()
                    .iter()
                    .position(|slide| slide.persist_id() == selected.persist_id())
                    .map(Position::new)
                    .ok_or(Error::Refused(Refusal::UncommittedSlideDependency))?;
                let bytes =
                    records::owner_record(&self.working, records::Owner::Slide(source_position))?;
                Some((bytes, source_position))
            },
            None => None,
        };
        let document_bytes = document.then(|| self.working.document.bytes().to_vec());
        let replacement = edit(records::View {
            slide: slide
                .as_ref()
                .map(|(bytes, _)| records::SlideRecord { bytes }),
            document: document_bytes.as_deref(),
        })?;
        if replacement.is_noop() {
            return Ok(false);
        }
        self.require_change_allowed()?;
        let mut rewrites = Vec::with_capacity(2);
        if let (Some(after), Some((before, position))) = (replacement.slide, slide) {
            rewrites.push(records::Rewrite {
                owner: records::Owner::Slide(position),
                before: Arc::from(before),
                after: Arc::from(after),
            });
        }
        let document_changed = replacement.document.is_some();
        if let (Some(after), Some(before)) = (replacement.document, document_bytes) {
            rewrites.push(records::Rewrite {
                owner: records::Owner::Document,
                before: Arc::from(before),
                after: Arc::from(after),
            });
        }
        if rewrites
            .iter()
            .all(|rewrite| rewrite.before == rewrite.after)
        {
            return Ok(false);
        }
        self.working = records::publish(&self.working, &rewrites)?;
        if document_changed {
            // Staged structure is rebased over the changed live document.
            self.live_document_changed = true;
        }
        self.formatting
            .push(FormattingChange::Element(records::ElementChange {
                rewrites,
            }));
        Ok(true)
    }

    fn require_change_allowed(&self) -> Result<()> {
        if self.source.has_signature {
            return Err(Error::Refused(Refusal::Signed));
        }
        Ok(())
    }

    fn order_fingerprint(&self) -> Result<String> {
        order_digest(&self.document).map(hex_digest)
    }

    /// Record a staged document-structure edit with the slide-order
    /// fingerprint its durable operation binds (record 0357b).
    fn push_document(&mut self, change: DocumentChange) -> Result<()> {
        let order = order_digest(&self.document)?;
        self.structural
            .push(StructuralChange::Document(change, order));
        Ok(())
    }
}

/// Successful immutable target and its reversible source-checked patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Published whole-package snapshot.
    #[must_use]
    pub(crate) const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible patch from the exact source to the committed artifact.
    #[must_use]
    pub(crate) const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Splits the commit into its target and patch.
    #[must_use]
    pub(crate) fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Reversible exact-source-checked whole-package slide-order patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Patch {
    before: Snapshot,
    after: Snapshot,
    changes: Vec<Change>,
    structural: Vec<StructuralChange>,
    visibility_changes: Vec<SlideVisibilityChange>,
    advance_changes: Vec<SlideAdvanceChange>,
    transition_visual_changes: Vec<SlideTransitionVisualChange>,
    text_changes: Vec<ShapeTextChange>,
    anchor_changes: Vec<ShapeAnchorChange>,
    media_path_changes: Vec<ExternalMediaPathChange>,
    media_playback_changes: Vec<ExternalMediaPlaybackChange>,
    formatting: Vec<FormattingChange>,
    formatting_first: bool,
    structural_before: StructuralArtifact,
    structural_after: StructuralArtifact,
}

/// The exact artifact that one side of a patch's structural operations is
/// bound to, in that patch's own direction.
///
/// Durable structural operations carry the SHA-256 of these artifacts as
/// `artifact_sha256` preconditions. When the artifact is one the patch already
/// retains, the digest is computed on demand from those retained bytes by
/// [`Patch::to_durable`] and memoized with the snapshot; commit computes none.
#[derive(Debug, Clone, PartialEq, Eq)]
enum StructuralArtifact {
    /// The patch's retained source artifact, [`Patch::before`].
    Before,
    /// The patch's retained target artifact, [`Patch::after`].
    After,
    /// The lowercase hexadecimal SHA-256 of an intermediate artifact between
    /// formatting and structural publication. The patch does not retain that
    /// artifact, so its digest was computed when the patch was built.
    Intermediate(String),
}

impl StructuralArtifact {
    /// Whether the committed source artifact is exactly the working artifact
    /// that structural publication started from. Without staged formatting the
    /// working snapshot shares the source allocation.
    fn retains_structural_source(source: &Snapshot, working: &Snapshot) -> bool {
        source.artifact.same_allocation(&working.artifact) || source.bytes() == working.bytes()
    }

    /// Binds the artifact that structural publication started from.
    fn structural_source(source: &Snapshot, working: &Snapshot) -> Self {
        if Self::retains_structural_source(source, working) {
            Self::Before
        } else {
            Self::intermediate(working)
        }
    }

    /// Captures the digest of a working artifact the patch will not retain.
    fn intermediate(working: &Snapshot) -> Self {
        Self::Intermediate(working.artifact_digest().to_owned())
    }

    /// The same artifact seen from the inverse patch, whose retained source
    /// and target artifacts are swapped.
    fn inverse(&self) -> Self {
        match self {
            Self::Before => Self::After,
            Self::After => Self::Before,
            Self::Intermediate(digest) => Self::Intermediate(digest.clone()),
        }
    }
}

impl Patch {
    /// Semantic moves represented by this patch.
    #[must_use]
    pub(crate) fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Slide hidden-state replacements composed into this root patch.
    #[must_use]
    pub(crate) fn slide_visibility_changes(&self) -> &[SlideVisibilityChange] {
        &self.visibility_changes
    }

    /// Slide advance replacements composed into this root patch.
    #[must_use]
    pub(crate) fn slide_advance_changes(&self) -> &[SlideAdvanceChange] {
        &self.advance_changes
    }

    /// Visual transition replacements composed into this root patch.
    #[must_use]
    pub(crate) fn slide_transition_visual_changes(&self) -> &[SlideTransitionVisualChange] {
        &self.transition_visual_changes
    }

    /// Speaker-notes text replacements composed into this root patch, in
    /// operation order.
    pub(crate) fn notes_text_changes(&self) -> impl Iterator<Item = &NotesTextChange> {
        self.formatting.iter().filter_map(|change| match change {
            FormattingChange::NotesText(change) => Some(change),
            _ => None,
        })
    }

    /// Shape-text replacements composed into this root patch.
    #[must_use]
    pub(crate) fn shape_text_changes(&self) -> &[ShapeTextChange] {
        &self.text_changes
    }

    /// Shape-anchor formatting replacements composed into this root patch.
    #[must_use]
    pub(crate) fn shape_anchor_changes(&self) -> &[ShapeAnchorChange] {
        &self.anchor_changes
    }

    /// External movie/audio path replacements composed into this root patch.
    #[must_use]
    pub(crate) fn external_media_path_changes(&self) -> &[ExternalMediaPathChange] {
        &self.media_path_changes
    }

    /// External-media playback replacements composed into this root patch.
    #[must_use]
    pub(crate) fn external_media_playback_changes(&self) -> &[ExternalMediaPlaybackChange] {
        &self.media_playback_changes
    }

    /// Whether this patch changes no artifact bytes.
    #[must_use]
    pub(crate) fn is_empty(&self) -> bool {
        self.before.bytes() == self.after.bytes()
    }

    /// Exact source artifact required for forward application.
    #[must_use]
    pub(crate) fn before(&self) -> &[u8] {
        self.before.bytes()
    }

    /// Exact committed target artifact.
    #[must_use]
    pub(crate) fn after(&self) -> &[u8] {
        self.after.bytes()
    }

    /// Applies only to the exact source artifact used to build this patch.
    ///
    /// # Errors
    ///
    /// Returns an error when `current` is not the exact source snapshot.
    pub(crate) fn apply(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.bytes() != self.before.bytes() {
            return Err(PackageError::InvalidFormat(
                "PPT slide-order patch source does not match its base artifact".into(),
            )
            .into());
        }
        Ok(self.after.clone())
    }

    /// An exact byte patch from `before` to `after` that retains no semantic
    /// operations (the result of a durable replay, record 0357a).
    pub(crate) fn exact(before: Snapshot, after: Snapshot) -> Self {
        Self {
            before,
            after,
            changes: Vec::new(),
            structural: Vec::new(),
            visibility_changes: Vec::new(),
            advance_changes: Vec::new(),
            transition_visual_changes: Vec::new(),
            text_changes: Vec::new(),
            anchor_changes: Vec::new(),
            media_path_changes: Vec::new(),
            media_playback_changes: Vec::new(),
            formatting: Vec::new(),
            formatting_first: true,
            structural_before: StructuralArtifact::Before,
            structural_after: StructuralArtifact::After,
        }
    }

    /// Refuses a patch whose operations a join cannot replay by value
    /// (record 0357b): slide insertions and removals shift every later
    /// position, element record edits replace persisted records the other
    /// patch may also rewrite, and an inverse patch that edits both
    /// formatting and structure resolves its formatting targets after its
    /// structure.
    fn require_joinable(&self) -> Result<()> {
        let refuse = |reason| Err(Error::Refused(Refusal::UnsupportedPatch { reason }));
        if self.structural.iter().any(|change| {
            matches!(
                change,
                StructuralChange::Remove(_) | StructuralChange::Insert(_)
            )
        }) {
            return refuse("slide insertions and removals cannot be joined");
        }
        if self
            .formatting
            .iter()
            .any(|change| matches!(change, FormattingChange::Element(_)))
        {
            return refuse("element record edits cannot be joined");
        }
        if !self.formatting_first && !self.formatting.is_empty() && !self.structural.is_empty() {
            return refuse("an inverse patch with formatting and structure cannot be joined");
        }
        Ok(())
    }

    /// The slide positions a patch permutes: the union of the inclusive
    /// ranges of its moves. Positions outside it are fixed points of every
    /// move.
    fn moved_positions(&self) -> BTreeSet<usize> {
        self.changes
            .iter()
            .flat_map(|change| {
                let start = change.from.get().min(change.destination.get());
                let end = change.from.get().max(change.destination.get());
                start..=end
            })
            .collect()
    }

    /// The slide-list positions a patch's structure depends on: its moved
    /// positions and the slides whose `fNonOutlineData` it sets. A
    /// position outside the moved set reads the same slide in every frame
    /// of the patch.
    fn slide_list_positions(&self) -> BTreeSet<usize> {
        let mut positions = self.moved_positions();
        positions.extend(self.structural.iter().filter_map(|change| match change {
            StructuralChange::Document(DocumentChange::NonOutlineData { position, .. }, _) => {
                Some(*position)
            },
            _ => None,
        }));
        positions
    }

    /// The semantic targets a patch writes, as three-way plans name them,
    /// plus its speaker-notes bodies.
    fn join_effects(&self) -> BTreeSet<String> {
        let mut effects = patch_effects(self);
        effects.extend(self.formatting.iter().filter_map(|change| match change {
            FormattingChange::NotesText(notes) => {
                Some(format!("slide:{}/notes-text", notes.slide.get()))
            },
            _ => None,
        }));
        effects
    }

    /// Stage this patch's formatting changes, by value, in `transaction`
    /// (whose structure is still the source's, so source positions are the
    /// projected positions).
    fn replay_formatting(&self, transaction: &mut Transaction) -> Result<()> {
        for change in &self.formatting {
            match change {
                FormattingChange::Text(text) => {
                    transaction.set_shape_text(text.target, text.after.clone())?;
                },
                FormattingChange::NotesText(notes) => {
                    transaction.set_notes_text(notes.slide, notes.after.clone())?;
                },
                FormattingChange::Anchor(anchor) => {
                    transaction.set_shape_anchor(anchor.target, anchor.after)?;
                },
                FormattingChange::SlideVisibility(visibility) => {
                    transaction.set_slide_hidden(visibility.target, visibility.after)?;
                },
                FormattingChange::SlideAdvance(advance) => {
                    transaction.set_slide_advance(advance.target, advance.after)?;
                },
                FormattingChange::SlideTransitionVisual(visual) => {
                    transaction.set_slide_transition_visual(visual.target, visual.after)?;
                },
                FormattingChange::MediaPath(media) => {
                    transaction.set_external_media_path(media.id, media.after.clone())?;
                },
                FormattingChange::MediaPlayback(media) => {
                    transaction.set_external_media_playback(media.id, media.after)?;
                },
                FormattingChange::Element(_) => {
                    return Err(Error::Refused(Refusal::UnsupportedPatch {
                        reason: "element record edits cannot be joined",
                    }));
                },
            }
        }
        Ok(())
    }

    /// Stage this patch's slide moves and document-structure edits, in
    /// order, in `transaction`.
    fn replay_structure(&self, transaction: &mut Transaction) -> Result<()> {
        for change in &self.structural {
            match change {
                StructuralChange::Move(move_change) => {
                    transaction.move_slide(move_change.from, move_change.destination)?;
                },
                StructuralChange::Document(document, _) => match document {
                    DocumentChange::Masters { order } => transaction.reorder_masters(order)?,
                    DocumentChange::InsertStyles { record, placement } => {
                        transaction.insert_custom_table_styles(record.clone(), *placement)?;
                    },
                    DocumentChange::RemoveStyles { .. } => {
                        transaction.remove_custom_table_styles()?;
                    },
                    DocumentChange::ReplaceStyles { after, .. } => {
                        transaction.replace_custom_table_styles(after.clone())?;
                    },
                    DocumentChange::MoveStyles { after, .. } => {
                        transaction.move_custom_table_styles(*after)?;
                    },
                    DocumentChange::NonOutlineData {
                        position, after, ..
                    } => {
                        transaction.set_slide_non_outline_data(Position::new(*position), *after)?
                    },
                },
                StructuralChange::Remove(_) | StructuralChange::Insert(_) => {
                    return Err(Error::Refused(Refusal::UnsupportedPatch {
                        reason: "slide insertions and removals cannot be joined",
                    }));
                },
            }
        }
        Ok(())
    }

    /// Joins two independently prepared patches over this patch's exact
    /// source into one patch (records 0357a, 0357b).
    ///
    /// Both patches must start from the same exact artifact. They may hold
    /// slide moves, master reorders, custom table-style edits,
    /// `fNonOutlineData` changes, and the formatting edits (shape and notes
    /// text, anchors, slide visibility, advance and visual transition,
    /// external media path and playback). The two must commute:
    ///
    /// - neither patch's slide-list positions (moved ranges and
    ///   `fNonOutlineData` slides) may meet the other's moved positions, so
    ///   every position reads the same slide in both patches;
    /// - no semantic target is written by both (two master reorders, two
    ///   custom table-style edits, the same shape's text, the same slide's
    ///   flag or notes, ...), since their order would decide the result.
    ///
    /// The joined patch stages both patches' formatting changes by value on
    /// the source (they name source slide positions), then `self`'s
    /// structure and `other`'s, in one native commit.
    ///
    /// # Errors
    ///
    /// [`Refusal::UnsupportedPatch`] for another source, an operation a join
    /// cannot replay ([`Self::require_joinable`]), or positions that meet a
    /// moved range; [`Refusal::JoinConflict`] for a target both write; the
    /// native commit's error otherwise.
    pub(crate) fn join(&self, other: &Self) -> Result<Self> {
        if self.before.bytes() != other.before.bytes() {
            return Err(Error::Refused(Refusal::UnsupportedPatch {
                reason: "joined patches start from different artifacts",
            }));
        }
        self.require_joinable()?;
        other.require_joinable()?;
        if !self
            .slide_list_positions()
            .is_disjoint(&other.moved_positions())
            || !other
                .slide_list_positions()
                .is_disjoint(&self.moved_positions())
        {
            return Err(Error::Refused(Refusal::UnsupportedPatch {
                reason: "joined slide moves overlap",
            }));
        }
        let left = self.join_effects();
        if let Some(target) = other
            .join_effects()
            .into_iter()
            .find(|target| left.contains(target))
        {
            return Err(Error::Refused(Refusal::JoinConflict { target }));
        }
        let mut transaction = self.before.edit()?;
        self.replay_formatting(&mut transaction)?;
        other.replay_formatting(&mut transaction)?;
        self.replay_structure(&mut transaction)?;
        other.replay_structure(&mut transaction)?;
        Ok(transaction.commit()?.patch)
    }

    /// Returns the exact-source-checked inverse patch.
    #[must_use]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            changes: self
                .changes
                .iter()
                .rev()
                .copied()
                .map(Change::inverse)
                .collect(),
            structural: self
                .structural
                .iter()
                .rev()
                .map(inverse_structural)
                .collect(),
            visibility_changes: self
                .visibility_changes
                .iter()
                .rev()
                .map(|change| SlideVisibilityChange {
                    position: change.position,
                    target: change.target,
                    slide_id: change.slide_id,
                    before: change.after,
                    after: change.before,
                })
                .collect(),
            advance_changes: self
                .advance_changes
                .iter()
                .rev()
                .map(|change| SlideAdvanceChange {
                    target: change.target,
                    slide_id: change.slide_id,
                    before: change.after,
                    after: change.before,
                })
                .collect(),
            transition_visual_changes: self
                .transition_visual_changes
                .iter()
                .rev()
                .map(|change| SlideTransitionVisualChange {
                    target: change.target,
                    slide_id: change.slide_id,
                    before: change.after,
                    after: change.before,
                })
                .collect(),
            text_changes: self
                .text_changes
                .iter()
                .rev()
                .map(|change| ShapeTextChange {
                    target: change.target,
                    before: change.after.clone(),
                    after: change.before.clone(),
                })
                .collect(),
            anchor_changes: self
                .anchor_changes
                .iter()
                .rev()
                .map(|change| ShapeAnchorChange {
                    target: change.target,
                    before: change.after,
                    after: change.before,
                })
                .collect(),
            media_path_changes: self
                .media_path_changes
                .iter()
                .rev()
                .map(|change| ExternalMediaPathChange {
                    id: change.id,
                    before: change.after.clone(),
                    after: change.before.clone(),
                })
                .collect(),
            media_playback_changes: self
                .media_playback_changes
                .iter()
                .rev()
                .map(|change| ExternalMediaPlaybackChange {
                    id: change.id,
                    before: change.after,
                    after: change.before,
                })
                .collect(),
            formatting: self
                .formatting
                .iter()
                .rev()
                .map(|change| match change {
                    FormattingChange::Text(text_change) => {
                        FormattingChange::Text(ShapeTextChange {
                            target: text_change.target,
                            before: text_change.after.clone(),
                            after: text_change.before.clone(),
                        })
                    },
                    FormattingChange::NotesText(notes_change) => {
                        FormattingChange::NotesText(NotesTextChange {
                            slide: notes_change.slide,
                            before: notes_change.after.clone(),
                            after: notes_change.before.clone(),
                        })
                    },
                    FormattingChange::Anchor(anchor_change) => {
                        FormattingChange::Anchor(ShapeAnchorChange {
                            target: anchor_change.target,
                            before: anchor_change.after,
                            after: anchor_change.before,
                        })
                    },
                    FormattingChange::SlideVisibility(visibility_change) => {
                        FormattingChange::SlideVisibility(SlideVisibilityChange {
                            position: visibility_change.position,
                            target: visibility_change.target,
                            slide_id: visibility_change.slide_id,
                            before: visibility_change.after,
                            after: visibility_change.before,
                        })
                    },
                    FormattingChange::SlideAdvance(advance_change) => {
                        FormattingChange::SlideAdvance(SlideAdvanceChange {
                            target: advance_change.target,
                            slide_id: advance_change.slide_id,
                            before: advance_change.after,
                            after: advance_change.before,
                        })
                    },
                    FormattingChange::SlideTransitionVisual(visual_change) => {
                        FormattingChange::SlideTransitionVisual(SlideTransitionVisualChange {
                            target: visual_change.target,
                            slide_id: visual_change.slide_id,
                            before: visual_change.after,
                            after: visual_change.before,
                        })
                    },
                    FormattingChange::MediaPath(media_change) => {
                        FormattingChange::MediaPath(ExternalMediaPathChange {
                            id: media_change.id,
                            before: media_change.after.clone(),
                            after: media_change.before.clone(),
                        })
                    },
                    FormattingChange::MediaPlayback(media_change) => {
                        FormattingChange::MediaPlayback(ExternalMediaPlaybackChange {
                            id: media_change.id,
                            before: media_change.after,
                            after: media_change.before,
                        })
                    },
                    FormattingChange::Element(element_change) => {
                        FormattingChange::Element(element_change.inverse())
                    },
                })
                .collect(),
            formatting_first: !self.formatting_first,
            structural_before: self.structural_after.inverse(),
            structural_after: self.structural_before.inverse(),
        }
    }

    /// Resolves one structural artifact binding to its SHA-256, computing a
    /// retained artifact's digest on first use.
    fn structural_artifact_digest<'a>(&'a self, artifact: &'a StructuralArtifact) -> &'a str {
        match artifact {
            StructuralArtifact::Before => self.before.artifact_digest(),
            StructuralArtifact::After => self.after.artifact_digest(),
            StructuralArtifact::Intermediate(digest) => digest,
        }
    }

    /// Converts this patch to the common bounded deterministic-JSON envelope.
    ///
    /// Operations use only semantic zero-based positions, exact artifact
    /// SHA-256, and a content-free order fingerprint; native slide and persist
    /// identifiers are never serialized.
    ///
    /// Structural operations bind the SHA-256 of the exact artifacts they
    /// start from and end at. Commit does not compute those digests: the
    /// first serialization that needs one computes it from the artifact bytes
    /// this patch retains, and every clone of that snapshot, including the
    /// inverse patch and a later edit of the committed snapshot, reuses it.
    ///
    /// # Errors
    ///
    /// Returns an error when caller-selected patch bounds cannot represent the
    /// semantic operations.
    pub(crate) fn to_durable(
        &self,
        limits: litchi_core::patch::PatchLimits,
    ) -> std::result::Result<
        litchi_core::patch::Patch<litchi_core::patch::Reversible>,
        litchi_core::patch::PatchError,
    > {
        use litchi_core::patch::{BlobBundle, ReversibleOperation};

        let mut forward_blobs = BlobBundle::new(limits.blobs());
        let mut reverse_blobs = BlobBundle::new(limits.blobs());
        let mut formatting_operations = Vec::new();
        for change in &self.formatting {
            let pair = match change {
                FormattingChange::NotesText(notes_change) => ReversibleOperation::new(
                    durable_notes_text_operation(
                        limits,
                        notes_change.slide,
                        &notes_change.before,
                        &notes_change.after,
                    )?,
                    durable_notes_text_operation(
                        limits,
                        notes_change.slide,
                        &notes_change.after,
                        &notes_change.before,
                    )?,
                ),
                FormattingChange::Text(text_change) => ReversibleOperation::new(
                    durable_text_operation(
                        limits,
                        text_change.target,
                        &text_change.before,
                        &text_change.after,
                    )?,
                    durable_text_operation(
                        limits,
                        text_change.target,
                        &text_change.after,
                        &text_change.before,
                    )?,
                ),
                FormattingChange::Anchor(anchor_change) => ReversibleOperation::new(
                    durable_anchor_operation(
                        limits,
                        anchor_change.target,
                        anchor_change.before,
                        anchor_change.after,
                    )?,
                    durable_anchor_operation(
                        limits,
                        anchor_change.target,
                        anchor_change.after,
                        anchor_change.before,
                    )?,
                ),
                FormattingChange::SlideVisibility(visibility_change) => ReversibleOperation::new(
                    durable_visibility_operation(
                        limits,
                        visibility_change.target,
                        visibility_change.before,
                        visibility_change.after,
                    )?,
                    durable_visibility_operation(
                        limits,
                        visibility_change.target,
                        visibility_change.after,
                        visibility_change.before,
                    )?,
                ),
                FormattingChange::SlideAdvance(advance_change) => ReversibleOperation::new(
                    durable_slide_advance_operation(
                        limits,
                        advance_change.target,
                        advance_change.before,
                        advance_change.after,
                    )?,
                    durable_slide_advance_operation(
                        limits,
                        advance_change.target,
                        advance_change.after,
                        advance_change.before,
                    )?,
                ),
                FormattingChange::SlideTransitionVisual(visual_change) => ReversibleOperation::new(
                    durable_slide_transition_visual_operation(
                        limits,
                        visual_change.target,
                        visual_change.before,
                        visual_change.after,
                    )?,
                    durable_slide_transition_visual_operation(
                        limits,
                        visual_change.target,
                        visual_change.after,
                        visual_change.before,
                    )?,
                ),
                FormattingChange::MediaPath(media_change) => ReversibleOperation::new(
                    durable_media_path_operation(
                        limits,
                        media_change.id,
                        media_change.before.as_deref(),
                        media_change.after.as_deref(),
                    )?,
                    durable_media_path_operation(
                        limits,
                        media_change.id,
                        media_change.after.as_deref(),
                        media_change.before.as_deref(),
                    )?,
                ),
                FormattingChange::MediaPlayback(media_change) => ReversibleOperation::new(
                    durable_media_playback_operation(
                        limits,
                        media_change.id,
                        media_change.before,
                        media_change.after,
                    )?,
                    durable_media_playback_operation(
                        limits,
                        media_change.id,
                        media_change.after,
                        media_change.before,
                    )?,
                ),
                FormattingChange::Element(element_change) => ReversibleOperation::new(
                    records::durable_operation(limits, element_change, &mut forward_blobs)?,
                    records::durable_operation(
                        limits,
                        &element_change.inverse(),
                        &mut reverse_blobs,
                    )?,
                ),
            };
            formatting_operations.push(pair);
        }
        let mut structural_operations = Vec::new();
        for structural_change in &self.structural {
            let (forward, inverse) = match structural_change {
                StructuralChange::Move(move_change) => (
                    durable_operation(
                        limits,
                        move_change.from,
                        move_change.destination,
                        self.structural_artifact_digest(&self.structural_before),
                        move_change.before_order,
                    )?,
                    durable_operation(
                        limits,
                        move_change.destination,
                        move_change.from,
                        self.structural_artifact_digest(&self.structural_after),
                        move_change.after_order,
                    )?,
                ),
                StructuralChange::Remove(list_change) => {
                    let payload = encode_payload(&normalized_payload(&list_change.payload))
                        .map_err(|_error| litchi_core::patch::PatchError::InvalidText {
                            field: "slide transfer payload",
                        })?;
                    let reverse_blob = reverse_blobs.insert(payload)?;
                    (
                        durable_remove_operation(
                            limits,
                            list_change,
                            self.structural_artifact_digest(&self.structural_before),
                            list_change.before_order,
                        )?,
                        durable_insert_operation(
                            limits,
                            list_change,
                            self.structural_artifact_digest(&self.structural_after),
                            list_change.after_order,
                            &reverse_blob.as_hex(),
                        )?,
                    )
                },
                StructuralChange::Insert(list_change) => {
                    let payload = encode_payload(&normalized_payload(&list_change.payload))
                        .map_err(|_error| litchi_core::patch::PatchError::InvalidText {
                            field: "slide transfer payload",
                        })?;
                    let forward_blob = forward_blobs.insert(payload)?;
                    (
                        durable_insert_operation(
                            limits,
                            list_change,
                            self.structural_artifact_digest(&self.structural_before),
                            list_change.before_order,
                            &forward_blob.as_hex(),
                        )?,
                        durable_remove_operation(
                            limits,
                            list_change,
                            self.structural_artifact_digest(&self.structural_after),
                            list_change.after_order,
                        )?,
                    )
                },
                // Master order, custom table styles and fNonOutlineData
                // (record 0357b): bound like the slide-list operations, to
                // the structural artifact and the slide order they run at.
                StructuralChange::Document(change, order) => (
                    durable_document_operation(
                        limits,
                        change,
                        self.structural_artifact_digest(&self.structural_before),
                        *order,
                        &mut forward_blobs,
                    )?,
                    durable_document_operation(
                        limits,
                        &change.inverse(),
                        self.structural_artifact_digest(&self.structural_after),
                        *order,
                        &mut reverse_blobs,
                    )?,
                ),
            };
            structural_operations.push(ReversibleOperation::new(forward, inverse));
        }
        let mut operations = Vec::with_capacity(
            formatting_operations
                .len()
                .saturating_add(structural_operations.len()),
        );
        if self.formatting_first {
            operations.extend(formatting_operations);
            operations.extend(structural_operations);
        } else {
            operations.extend(structural_operations);
            operations.extend(formatting_operations);
        }
        litchi_core::patch::Patch::<litchi_core::patch::Reversible>::new(
            limits,
            "litchi-ppt",
            operations,
            forward_blobs,
            reverse_blobs,
        )
    }
}

fn durable_operation(
    limits: litchi_core::patch::PatchLimits,
    from: Position,
    destination: Position,
    artifact: &str,
    order: [u8; 32],
) -> std::result::Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "artifact_sha256".to_string(),
        serde_json::Value::String(artifact.to_string()),
    );
    preconditions.insert(
        "order_sha256".to_string(),
        serde_json::Value::String(hex_digest(order)),
    );
    let destination_value = u64::try_from(destination.get())
        .map(serde_json::Value::from)
        .map_err(|_error| litchi_core::patch::PatchError::InvalidText {
            field: "slide destination",
        })?;
    litchi_core::patch::PatchOperation::new(
        limits,
        "slide-order.move",
        format!("slide:{}", from.get()),
        preconditions,
        destination_value,
    )
}

fn durable_visibility_operation(
    limits: litchi_core::patch::PatchLimits,
    target: Position,
    before: bool,
    after: bool,
) -> std::result::Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    let preconditions = BTreeMap::from([("hidden".to_string(), serde_json::Value::Bool(before))]);
    litchi_core::patch::PatchOperation::new(
        limits,
        "presentation.slide-hidden.set",
        format!("slide:{}", target.get()),
        preconditions,
        serde_json::Value::Bool(after),
    )
}

fn parse_target(target: &str) -> Result<Position> {
    let value = target
        .strip_prefix("slide:")
        .filter(|value| {
            !value.is_empty()
                && value.bytes().all(|byte| byte.is_ascii_digit())
                && (value == &"0" || !value.starts_with('0'))
        })
        .ok_or_else(|| invalid_durable_patch("invalid slide-order target"))?;
    value
        .parse::<usize>()
        .map(Position::new)
        .map_err(|_error| invalid_durable_patch("slide position exceeds this platform"))
}

fn invalid_durable_patch(message: &str) -> Error {
    PackageError::InvalidFormat(format!("invalid PPT durable patch: {message}")).into()
}

fn artifact_hash(bytes: &[u8]) -> String {
    litchi_core::patch::BlobId::of(bytes).as_hex()
}

fn durable_text_operation(
    limits: litchi_core::patch::PatchLimits,
    target: crate::text_edit::Target,
    before: &str,
    after: &str,
) -> std::result::Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "text_sha256".to_string(),
        serde_json::Value::String(artifact_hash(before.as_bytes())),
    );
    litchi_core::patch::PatchOperation::new(
        limits,
        "presentation.shape-text.set",
        format!(
            "slide:{}/shape:{}",
            target.slide().get(),
            target.shape().get()
        ),
        preconditions,
        serde_json::Value::String(after.to_string()),
    )
}

fn durable_notes_text_operation(
    limits: litchi_core::patch::PatchLimits,
    slide: Position,
    before: &str,
    after: &str,
) -> std::result::Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "text_sha256".to_string(),
        serde_json::Value::String(artifact_hash(before.as_bytes())),
    );
    litchi_core::patch::PatchOperation::new(
        limits,
        "presentation.notes-text.set",
        format!("slide:{}", slide.get()),
        preconditions,
        serde_json::Value::String(after.to_string()),
    )
}

fn durable_anchor_operation(
    limits: litchi_core::patch::PatchLimits,
    target: crate::text_edit::Target,
    before: crate::Anchor,
    after: crate::Anchor,
) -> std::result::Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "anchor_sha256".to_string(),
        serde_json::Value::String(artifact_hash(&before.to_bytes())),
    );
    let encoding = match after.encoding() {
        crate::client_anchor::Encoding::Small => "small",
        crate::client_anchor::Encoding::Full => "full",
    };
    let value = serde_json::json!({
        "encoding": encoding,
        "left": after.left(),
        "top": after.top(),
        "right": after.right(),
        "bottom": after.bottom(),
    });
    litchi_core::patch::PatchOperation::new(
        limits,
        "presentation.shape-anchor.set",
        format!(
            "slide:{}/shape:{}",
            target.slide().get(),
            target.shape().get()
        ),
        preconditions,
        value,
    )
}

fn durable_slide_advance_operation(
    limits: litchi_core::patch::PatchLimits,
    target: Position,
    before: SlideAdvance,
    after: SlideAdvance,
) -> std::result::Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    let preconditions = BTreeMap::from([(
        "advance_sha256".to_string(),
        serde_json::Value::String(artifact_hash(&slide_advance_bytes(before))),
    )]);
    litchi_core::patch::PatchOperation::new(
        limits,
        "presentation.slide-advance.set",
        format!("slide:{}", target.get()),
        preconditions,
        slide_advance_value(after),
    )
}

fn slide_advance_bytes(advance: SlideAdvance) -> [u8; 6] {
    let mut bytes = [0_u8; 6];
    bytes[0] = u8::from(advance.manual);
    bytes[1] = u8::from(advance.automatic);
    bytes[2..6].copy_from_slice(&advance.delay_ms.to_le_bytes());
    bytes
}

fn slide_advance_value(advance: SlideAdvance) -> serde_json::Value {
    serde_json::json!({
        "automatic": advance.automatic,
        "delay_ms": advance.delay_ms,
        "manual": advance.manual,
    })
}

fn parse_slide_advance_value(value: &serde_json::Value) -> Result<SlideAdvance> {
    let object = value
        .as_object()
        .filter(|object| object.len() == 3)
        .ok_or_else(|| invalid_durable_patch("slide-advance value must have three fields"))?;
    let manual = object
        .get("manual")
        .and_then(serde_json::Value::as_bool)
        .ok_or_else(|| invalid_durable_patch("slide-advance manual flag is invalid"))?;
    let automatic = object
        .get("automatic")
        .and_then(serde_json::Value::as_bool)
        .ok_or_else(|| invalid_durable_patch("slide-advance automatic flag is invalid"))?;
    let delay_ms = object
        .get("delay_ms")
        .and_then(serde_json::Value::as_u64)
        .and_then(|delay| u32::try_from(delay).ok())
        .ok_or_else(|| invalid_durable_patch("slide-advance delay is invalid"))?;
    SlideAdvance::new(manual, automatic, delay_ms)
}

fn durable_slide_transition_visual_operation(
    limits: litchi_core::patch::PatchLimits,
    target: Position,
    before: SlideTransitionVisual,
    after: SlideTransitionVisual,
) -> std::result::Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    let preconditions = BTreeMap::from([(
        "transition_visual_sha256".to_string(),
        serde_json::Value::String(artifact_hash(&transition_visual_bytes(before))),
    )]);
    litchi_core::patch::PatchOperation::new(
        limits,
        "presentation.slide-transition-visual.set",
        format!("slide:{}", target.get()),
        preconditions,
        transition_visual_value(after),
    )
}

fn transition_visual_bytes(visual: SlideTransitionVisual) -> [u8; 3] {
    visual.wire
}

fn transition_visual_value(visual: SlideTransitionVisual) -> serde_json::Value {
    let [effect_direction, effect_type, speed] = transition_visual_bytes(visual);
    serde_json::json!({
        "effect_direction": effect_direction,
        "effect_type": effect_type,
        "speed": speed,
    })
}

fn parse_transition_visual_value(value: &serde_json::Value) -> Result<SlideTransitionVisual> {
    let object = value
        .as_object()
        .filter(|object| object.len() == 3)
        .ok_or_else(|| invalid_durable_patch("transition-visual value must have three fields"))?;
    let field = |name| {
        object
            .get(name)
            .and_then(serde_json::Value::as_u64)
            .and_then(|number| u8::try_from(number).ok())
            .ok_or_else(|| invalid_durable_patch("transition-visual field is invalid"))
    };
    let bytes = [
        field("effect_direction")?,
        field("effect_type")?,
        field("speed")?,
    ];
    slide_transition_visual_from_bytes(bytes)
        .ok_or_else(|| invalid_durable_patch("transition-visual mapping is noncanonical"))
}

fn durable_media_path_operation(
    limits: litchi_core::patch::PatchLimits,
    id: u32,
    before: Option<&str>,
    after: Option<&str>,
) -> std::result::Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    let preconditions = BTreeMap::from([(
        "path_sha256".to_string(),
        serde_json::Value::String(media_path_digest(before)),
    )]);
    litchi_core::patch::PatchOperation::new(
        limits,
        "presentation.external-media-path.set",
        format!("media:{id}"),
        preconditions,
        after.map_or(serde_json::Value::Null, |path| {
            serde_json::Value::String(path.to_string())
        }),
    )
}

fn durable_media_playback_operation(
    limits: litchi_core::patch::PatchLimits,
    id: u32,
    before: crate::external_media::Playback,
    after: crate::external_media::Playback,
) -> std::result::Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    let preconditions = BTreeMap::from([(
        "playback_sha256".to_string(),
        serde_json::Value::String(playback_digest(before)),
    )]);
    litchi_core::patch::PatchOperation::new(
        limits,
        "presentation.external-media-playback.set",
        format!("media:{id}"),
        preconditions,
        playback_value(after),
    )
}

fn durable_remove_operation(
    limits: litchi_core::patch::PatchLimits,
    change: &ListChange,
    artifact: &str,
    order: [u8; 32],
) -> std::result::Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    let mut preconditions = structural_preconditions(artifact, order);
    preconditions.insert(
        "slide_sha256".to_string(),
        serde_json::Value::String(artifact_hash(&change.payload.record)),
    );
    litchi_core::patch::PatchOperation::new(
        limits,
        "slide-order.remove",
        format!("slide:{}", change.position.get()),
        preconditions,
        serde_json::Value::Null,
    )
}

fn durable_insert_operation(
    limits: litchi_core::patch::PatchLimits,
    change: &ListChange,
    artifact: &str,
    order: [u8; 32],
    blob: &str,
) -> std::result::Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    let mut preconditions = structural_preconditions(artifact, order);
    preconditions.insert(
        "master_id".to_string(),
        serde_json::Value::from(u64::from(change.payload.master_id)),
    );
    litchi_core::patch::PatchOperation::new(
        limits,
        "slide-order.insert",
        format!("slide:{}", change.position.get()),
        preconditions,
        serde_json::Value::String(blob.to_string()),
    )
}

/// The durable name of a custom table-style placement.
const fn placement_name(placement: document_structure::CustomTableStylesPlacement) -> &'static str {
    match placement {
        document_structure::CustomTableStylesPlacement::BeforeEndDocument => "before-end-document",
        document_structure::CustomTableStylesPlacement::AfterEndDocument => "after-end-document",
    }
}

fn parse_placement(
    value: Option<&serde_json::Value>,
) -> Result<document_structure::CustomTableStylesPlacement> {
    match value.and_then(serde_json::Value::as_str) {
        Some("before-end-document") => {
            Ok(document_structure::CustomTableStylesPlacement::BeforeEndDocument)
        },
        Some("after-end-document") => {
            Ok(document_structure::CustomTableStylesPlacement::AfterEndDocument)
        },
        _ => Err(invalid_durable_patch(
            "custom table-style placement is invalid",
        )),
    }
}

/// One durable document-structure operation (record 0357b). Every one is
/// bound to the structural artifact and the slide order it runs at, like
/// the slide-list operations, and to the state it replaces where the change
/// carries it: the SHA-256 of the custom table-style atom it removes or
/// replaces, the placement it moves from, and the `fNonOutlineData` value it
/// overwrites. Atoms travel as blobs. Master order is a permutation of
/// zero-based master positions; no native identifier is serialized.
fn durable_document_operation(
    limits: litchi_core::patch::PatchLimits,
    change: &DocumentChange,
    artifact: &str,
    order: [u8; 32],
    blobs: &mut litchi_core::patch::BlobBundle,
) -> std::result::Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    use serde_json::Value;

    let encode = |record: &crate::Record| {
        document_structure::encode_record(record, false).map_err(|_error| {
            litchi_core::patch::PatchError::InvalidText {
                field: "custom table-style atom",
            }
        })
    };
    let mut preconditions = structural_preconditions(artifact, order);
    let document = "document".to_owned();
    let (op, target, value) = match change {
        DocumentChange::Masters { order: positions } => {
            let positions = positions
                .iter()
                .map(|position| u64::try_from(*position).map(Value::from))
                .collect::<std::result::Result<Vec<_>, _>>()
                .map_err(|_error| litchi_core::patch::PatchError::InvalidText {
                    field: "master position",
                })?;
            (
                "document-structure.master-order.set",
                document,
                Value::Array(positions),
            )
        },
        DocumentChange::InsertStyles { record, placement } => {
            let blob = blobs.insert(encode(record)?)?;
            (
                "document-structure.custom-table-styles.insert",
                document,
                serde_json::json!({
                    "placement": placement_name(*placement),
                    "record": blob.as_hex(),
                }),
            )
        },
        DocumentChange::RemoveStyles { record, .. } => {
            preconditions.insert(
                "styles_sha256".to_owned(),
                Value::String(artifact_hash(&encode(record)?)),
            );
            (
                "document-structure.custom-table-styles.remove",
                document,
                Value::Null,
            )
        },
        DocumentChange::ReplaceStyles { before, after } => {
            preconditions.insert(
                "styles_sha256".to_owned(),
                Value::String(artifact_hash(&encode(before)?)),
            );
            let blob = blobs.insert(encode(after)?)?;
            (
                "document-structure.custom-table-styles.replace",
                document,
                Value::String(blob.as_hex()),
            )
        },
        DocumentChange::MoveStyles { before, after } => {
            preconditions.insert(
                "placement".to_owned(),
                Value::String(placement_name(*before).to_owned()),
            );
            (
                "document-structure.custom-table-styles.move",
                document,
                Value::String(placement_name(*after).to_owned()),
            )
        },
        DocumentChange::NonOutlineData {
            position,
            before,
            after,
        } => {
            preconditions.insert("non_outline_data".to_owned(), Value::Bool(*before));
            (
                "document-structure.slide-non-outline-data.set",
                format!("slide:{position}"),
                Value::Bool(*after),
            )
        },
    };
    litchi_core::patch::PatchOperation::new(limits, op, target, preconditions, value)
}

/// Replay one durable document-structure operation (record 0357b) in the
/// structural transaction, after its artifact and slide-order
/// preconditions were checked.
fn apply_durable_document(
    transaction: &mut Transaction,
    operation: &litchi_core::patch::PatchOperation,
    blobs: &litchi_core::patch::BlobBundle,
) -> Result<()> {
    let mismatch = || {
        Error::from(PackageError::InvalidFormat(
            "PPT durable document-structure precondition does not match".into(),
        ))
    };
    let precondition = |name: &str| operation.preconditions.get(name);
    let styles_digest = || {
        precondition("styles_sha256")
            .and_then(serde_json::Value::as_str)
            .ok_or_else(|| invalid_durable_patch("missing custom table-style precondition"))
    };
    let check_styles = |record: &crate::Record, expected: &str| -> Result<()> {
        let bytes = document_structure::encode_record(record, false)?;
        if artifact_hash(&bytes) == expected {
            Ok(())
        } else {
            Err(mismatch())
        }
    };
    let styles_record = |value: Option<&serde_json::Value>| -> Result<crate::Record> {
        let blob = value
            .and_then(serde_json::Value::as_str)
            .and_then(|hex| blob_by_hex(blobs, hex))
            .ok_or_else(|| invalid_durable_patch("custom table-style blob is missing"))?;
        let (record, consumed) = crate::Record::parse(blob, 0, &SourceLimits::default())
            .map_err(|_error| invalid_durable_patch("custom table-style blob is not a record"))?;
        if consumed != blob.len() {
            return Err(invalid_durable_patch(
                "custom table-style blob is not exactly one record",
            ));
        }
        Ok(record)
    };
    let count = operation.preconditions.len();
    let on_document = operation.target == "document";
    let value = &operation.value;
    match operation.op.as_str() {
        "document-structure.master-order.set" if count == 2 && on_document => {
            let order = value
                .as_array()
                .and_then(|positions| {
                    positions
                        .iter()
                        .map(|position| {
                            position
                                .as_u64()
                                .and_then(|position| usize::try_from(position).ok())
                        })
                        .collect::<Option<Vec<_>>>()
                })
                .ok_or_else(|| invalid_durable_patch("master order must be positions"))?;
            transaction.reorder_masters(&order)
        },
        "document-structure.custom-table-styles.insert" if count == 2 && on_document => {
            let object = value
                .as_object()
                .filter(|object| object.len() == 2)
                .ok_or_else(|| invalid_durable_patch("custom table-style insertion is invalid"))?;
            let placement = parse_placement(object.get("placement"))?;
            let record = styles_record(object.get("record"))?;
            transaction.insert_custom_table_styles(record, placement)
        },
        "document-structure.custom-table-styles.remove"
            if count == 3 && on_document && value.is_null() =>
        {
            let expected = styles_digest()?;
            let removed = transaction.remove_custom_table_styles()?;
            check_styles(&removed, expected)
        },
        "document-structure.custom-table-styles.replace" if count == 3 && on_document => {
            let expected = styles_digest()?;
            let record = styles_record(Some(value))?;
            let before = transaction.replace_custom_table_styles(record)?;
            check_styles(&before, expected)
        },
        "document-structure.custom-table-styles.move" if count == 3 && on_document => {
            let before = parse_placement(precondition("placement"))?;
            let after = parse_placement(Some(value))?;
            if transaction.require_custom_table_styles()? != before {
                return Err(mismatch());
            }
            transaction.move_custom_table_styles(after)
        },
        "document-structure.slide-non-outline-data.set" if count == 3 => {
            let position = parse_target(&operation.target)?;
            let before = precondition("non_outline_data")
                .and_then(serde_json::Value::as_bool)
                .ok_or_else(|| invalid_durable_patch("missing fNonOutlineData precondition"))?;
            let after = value
                .as_bool()
                .ok_or_else(|| invalid_durable_patch("fNonOutlineData value must be a flag"))?;
            transaction.require_position(position)?;
            let current = transaction.document.slides()?[position.get()].flags() & (1 << 2) != 0;
            if current != before {
                return Err(mismatch());
            }
            transaction.set_slide_non_outline_data(position, after)
        },
        _ => Err(invalid_durable_patch(
            "unsupported document-structure operation vocabulary",
        )),
    }
}

fn structural_preconditions(
    artifact: &str,
    order: [u8; 32],
) -> BTreeMap<String, serde_json::Value> {
    BTreeMap::from([
        (
            "artifact_sha256".to_string(),
            serde_json::Value::String(artifact.to_string()),
        ),
        (
            "order_sha256".to_string(),
            serde_json::Value::String(hex_digest(order)),
        ),
    ])
}

fn apply_durable_text(
    snapshot: &Snapshot,
    operation: &litchi_core::patch::PatchOperation,
) -> Result<Snapshot> {
    if operation.preconditions.len() != 1 {
        return Err(invalid_durable_patch(
            "shape-text operation has unexpected preconditions",
        ));
    }
    let target = parse_shape_target(&operation.target)?;
    let replacement = operation
        .value
        .as_str()
        .ok_or_else(|| invalid_durable_patch("shape-text value must be a string"))?;
    let expected = operation
        .preconditions
        .get("text_sha256")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| invalid_durable_patch("missing shape-text precondition"))?;
    if snapshot.has_signature {
        let current = crate::text_edit::inspect_shape_text_readonly(
            snapshot.bytes(),
            target,
            snapshot.open_options(),
        )?;
        if artifact_hash(current.as_bytes()) != expected {
            return Err(PackageError::InvalidFormat(
                "PPT durable shape-text semantic precondition does not match".into(),
            )
            .into());
        }
        if current == replacement {
            return Ok(snapshot.clone());
        }
        return Err(Error::Refused(Refusal::Signed));
    }
    let text_snapshot = crate::text_edit::Snapshot::from_bytes(snapshot.bytes().to_vec())?;
    let mut edit = match text_snapshot.edit_text(target) {
        Ok(edit) => edit,
        Err(crate::text_edit::Error::Refused(crate::text_edit::Refusal::NoTextAtom)) => {
            // Outline-backed placeholder text lives in the live document;
            // the root transaction owns that publication.
            let current = crate::text_edit::inspect_shape_text_readonly(
                snapshot.bytes(),
                target,
                snapshot.open_options(),
            )?;
            if artifact_hash(current.as_bytes()) != expected {
                return Err(PackageError::InvalidFormat(
                    "PPT durable shape-text semantic precondition does not match".into(),
                )
                .into());
            }
            let mut transaction = snapshot.edit()?;
            transaction.set_shape_text(target, replacement)?;
            return Ok(transaction.commit()?.snapshot);
        },
        Err(error) => return Err(error.into()),
    };
    if artifact_hash(edit.text().as_bytes()) != expected {
        return Err(PackageError::InvalidFormat(
            "PPT durable shape-text semantic precondition does not match".into(),
        )
        .into());
    }
    edit.set_text(replacement)?;
    let committed = edit.commit()?;
    Snapshot::from_bytes_with_options(
        committed.snapshot().bytes().to_vec(),
        snapshot.open_options(),
    )
}

fn apply_durable_notes_text(
    snapshot: &Snapshot,
    operation: &litchi_core::patch::PatchOperation,
) -> Result<Snapshot> {
    if operation.preconditions.len() != 1 {
        return Err(invalid_durable_patch(
            "notes-text operation has unexpected preconditions",
        ));
    }
    let slide = parse_target(&operation.target)?;
    let replacement = operation
        .value
        .as_str()
        .ok_or_else(|| invalid_durable_patch("notes-text value must be a string"))?;
    let expected = operation
        .preconditions
        .get("text_sha256")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| invalid_durable_patch("missing notes-text precondition"))?;
    let current = crate::text_edit::inspect_notes_text_readonly(
        snapshot.bytes(),
        slide,
        snapshot.open_options(),
    )?;
    if artifact_hash(current.as_bytes()) != expected {
        return Err(PackageError::InvalidFormat(
            "PPT durable notes-text semantic precondition does not match".into(),
        )
        .into());
    }
    let mut transaction = snapshot.edit()?;
    transaction.set_notes_text(slide, replacement)?;
    Ok(transaction.commit()?.snapshot)
}

fn is_formatting_operation(operation: &str) -> bool {
    matches!(
        operation,
        "presentation.shape-text.set"
            | "presentation.notes-text.set"
            | "presentation.shape-anchor.set"
            | "presentation.slide-hidden.set"
            | "presentation.slide-advance.set"
            | "presentation.slide-transition-visual.set"
            | "presentation.external-media-path.set"
            | "presentation.external-media-playback.set"
            | "presentation.element-records.set"
    )
}

fn apply_durable_media_path(
    snapshot: &Snapshot,
    operation: &litchi_core::patch::PatchOperation,
) -> Result<Snapshot> {
    if operation.preconditions.len() != 1 {
        return Err(invalid_durable_patch(
            "external-media path operation has unexpected preconditions",
        ));
    }
    let id = parse_media_target(&operation.target)?;
    let replacement = match &operation.value {
        serde_json::Value::Null => None,
        serde_json::Value::String(path) => Some(path.clone()),
        serde_json::Value::Bool(_)
        | serde_json::Value::Number(_)
        | serde_json::Value::Array(_)
        | serde_json::Value::Object(_) => {
            return Err(invalid_durable_patch(
                "external-media path value must be a string or null",
            ));
        },
    };
    let expected = operation
        .preconditions
        .get("path_sha256")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| invalid_durable_patch("missing external-media path precondition"))?;
    let current = snapshot.external_media_path(id)?;
    if media_path_digest(current.as_deref()) != expected {
        return Err(PackageError::InvalidFormat(
            "PPT durable external-media path precondition does not match".into(),
        )
        .into());
    }
    if snapshot.has_signature {
        if current == replacement {
            return Ok(snapshot.clone());
        }
        return Err(Error::Refused(Refusal::Signed));
    }
    let mut transaction = snapshot.edit()?;
    transaction.set_external_media_path(id, replacement)?;
    transaction.commit().map(|commit| commit.snapshot)
}

fn apply_durable_media_playback(
    snapshot: &Snapshot,
    operation: &litchi_core::patch::PatchOperation,
) -> Result<Snapshot> {
    if operation.preconditions.len() != 1 {
        return Err(invalid_durable_patch(
            "external-media playback operation has unexpected preconditions",
        ));
    }
    let id = parse_media_target(&operation.target)?;
    let replacement = parse_playback_value(&operation.value)?;
    let expected = operation
        .preconditions
        .get("playback_sha256")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| invalid_durable_patch("missing external-media playback precondition"))?;
    let current = snapshot.external_media_playback(id)?;
    if playback_digest(current) != expected {
        return Err(PackageError::InvalidFormat(
            "PPT durable external-media playback precondition does not match".into(),
        )
        .into());
    }
    if snapshot.has_signature {
        if current == replacement {
            return Ok(snapshot.clone());
        }
        return Err(Error::Refused(Refusal::Signed));
    }
    let mut transaction = snapshot.edit()?;
    transaction.set_external_media_playback(id, replacement)?;
    transaction.commit().map(|commit| commit.snapshot)
}

fn parse_media_target(target: &str) -> Result<u32> {
    let value = target
        .strip_prefix("media:")
        .filter(|value| {
            !value.is_empty()
                && value.bytes().all(|byte| byte.is_ascii_digit())
                && (value == &"0" || !value.starts_with('0'))
        })
        .ok_or_else(|| invalid_durable_patch("invalid external-media target"))?;
    value
        .parse::<u32>()
        .map_err(|_error| invalid_durable_patch("external-media ID exceeds u32"))
}

fn media_path_digest(path: Option<&str>) -> String {
    let mut bytes = Vec::with_capacity(path.map_or(1, |value| value.len().saturating_add(1)));
    match path {
        Some(value) => {
            bytes.push(1);
            bytes.extend_from_slice(value.as_bytes());
        },
        None => bytes.push(0),
    }
    artifact_hash(&bytes)
}

fn playback_digest(playback: crate::external_media::Playback) -> String {
    artifact_hash(&[
        u8::from(playback.loop_playback),
        u8::from(playback.rewind_after_playing),
        u8::from(playback.narration),
    ])
}

fn playback_value(playback: crate::external_media::Playback) -> serde_json::Value {
    serde_json::json!({
        "loop": playback.loop_playback,
        "narration": playback.narration,
        "rewind": playback.rewind_after_playing,
    })
}

fn parse_playback_value(value: &serde_json::Value) -> Result<crate::external_media::Playback> {
    let object = value
        .as_object()
        .filter(|object| object.len() == 3)
        .ok_or_else(|| {
            invalid_durable_patch("external-media playback value must have exactly three fields")
        })?;
    let flag = |name: &str| {
        object
            .get(name)
            .and_then(serde_json::Value::as_bool)
            .ok_or_else(|| invalid_durable_patch("external-media playback flag is invalid"))
    };
    Ok(crate::external_media::Playback::new(
        flag("loop")?,
        flag("rewind")?,
        flag("narration")?,
    ))
}

fn apply_durable_anchor(
    snapshot: &Snapshot,
    operation: &litchi_core::patch::PatchOperation,
) -> Result<Snapshot> {
    if operation.preconditions.len() != 1 {
        return Err(invalid_durable_patch(
            "shape-anchor operation has unexpected preconditions",
        ));
    }
    let target = parse_shape_target(&operation.target)?;
    let expected = operation
        .preconditions
        .get("anchor_sha256")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| invalid_durable_patch("missing shape-anchor precondition"))?;
    let replacement = parse_anchor_value(&operation.value)?;
    if snapshot.has_signature {
        let current = crate::text_edit::inspect_shape_anchor_readonly(
            snapshot.bytes(),
            target,
            snapshot.open_options(),
        )?;
        if artifact_hash(&current.to_bytes()) != expected {
            return Err(PackageError::InvalidFormat(
                "PPT durable shape-anchor semantic precondition does not match".into(),
            )
            .into());
        }
        if current == replacement {
            return Ok(snapshot.clone());
        }
        return Err(Error::Refused(Refusal::Signed));
    }
    let current = crate::text_edit::inspect_shape_anchor(snapshot.bytes(), target)?;
    if artifact_hash(&current.to_bytes()) != expected {
        return Err(PackageError::InvalidFormat(
            "PPT durable shape-anchor semantic precondition does not match".into(),
        )
        .into());
    }
    let rewritten = crate::text_edit::replace_shape_anchor(snapshot.bytes(), target, replacement)?;
    Snapshot::from_bytes_with_options(rewritten.bytes, snapshot.open_options())
}

fn apply_durable_slide_visibility(
    snapshot: &Snapshot,
    operation: &litchi_core::patch::PatchOperation,
) -> Result<Snapshot> {
    if operation.preconditions.len() != 1 {
        return Err(invalid_durable_patch(
            "slide-hidden operation has unexpected preconditions",
        ));
    }
    let target = parse_target(&operation.target)?;
    let expected = operation
        .preconditions
        .get("hidden")
        .and_then(serde_json::Value::as_bool)
        .ok_or_else(|| invalid_durable_patch("missing slide-hidden precondition"))?;
    if snapshot.slide_hidden(target)? != expected {
        return Err(PackageError::InvalidFormat(
            "PPT durable slide-hidden semantic precondition does not match".into(),
        )
        .into());
    }
    let replacement = operation
        .value
        .as_bool()
        .ok_or_else(|| invalid_durable_patch("slide-hidden value must be boolean"))?;
    if snapshot.has_signature {
        if snapshot.slide_hidden(target)? == replacement {
            return Ok(snapshot.clone());
        }
        return Err(Error::Refused(Refusal::Signed));
    }
    replace_slide_hidden(snapshot, target, replacement)
}

fn apply_durable_slide_advance(
    snapshot: &Snapshot,
    operation: &litchi_core::patch::PatchOperation,
) -> Result<Snapshot> {
    if operation.preconditions.len() != 1 {
        return Err(invalid_durable_patch(
            "slide-advance operation has unexpected preconditions",
        ));
    }
    let target = parse_target(&operation.target)?;
    let expected = operation
        .preconditions
        .get("advance_sha256")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| invalid_durable_patch("missing slide-advance precondition"))?;
    let current = snapshot.slide_advance(target)?;
    if artifact_hash(&slide_advance_bytes(current)) != expected {
        return Err(PackageError::InvalidFormat(
            "PPT durable slide-advance semantic precondition does not match".into(),
        )
        .into());
    }
    let replacement = parse_slide_advance_value(&operation.value)?;
    if snapshot.has_signature {
        if current == replacement {
            return Ok(snapshot.clone());
        }
        return Err(Error::Refused(Refusal::Signed));
    }
    replace_slide_advance(snapshot, target, replacement)
}

fn apply_durable_slide_transition_visual(
    snapshot: &Snapshot,
    operation: &litchi_core::patch::PatchOperation,
) -> Result<Snapshot> {
    if operation.preconditions.len() != 1 {
        return Err(invalid_durable_patch(
            "slide-transition-visual operation has unexpected preconditions",
        ));
    }
    let target = parse_target(&operation.target)?;
    let expected = operation
        .preconditions
        .get("transition_visual_sha256")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| invalid_durable_patch("missing transition-visual precondition"))?;
    let current = snapshot.slide_transition_visual(target)?;
    if artifact_hash(&transition_visual_bytes(current)) != expected {
        return Err(PackageError::InvalidFormat(
            "PPT durable transition-visual semantic precondition does not match".into(),
        )
        .into());
    }
    let replacement = parse_transition_visual_value(&operation.value)?;
    if snapshot.has_signature {
        if current == replacement {
            return Ok(snapshot.clone());
        }
        return Err(Error::Refused(Refusal::Signed));
    }
    replace_slide_transition_visual(snapshot, target, replacement)
}

fn parse_anchor_value(value: &serde_json::Value) -> Result<crate::Anchor> {
    let object = value
        .as_object()
        .filter(|object| object.len() == 5)
        .ok_or_else(|| invalid_durable_patch("shape-anchor value must have exactly five fields"))?;
    let coordinate = |name: &str| {
        object
            .get(name)
            .and_then(serde_json::Value::as_i64)
            .and_then(|coordinate_value| i32::try_from(coordinate_value).ok())
            .ok_or_else(|| invalid_durable_patch("shape-anchor coordinate is invalid"))
    };
    let left = coordinate("left")?;
    let top = coordinate("top")?;
    let right = coordinate("right")?;
    let bottom = coordinate("bottom")?;
    match object.get("encoding").and_then(serde_json::Value::as_str) {
        Some("small") => crate::Anchor::small(
            i16::try_from(left)
                .map_err(|_error| invalid_durable_patch("small anchor left is out of range"))?,
            i16::try_from(top)
                .map_err(|_error| invalid_durable_patch("small anchor top is out of range"))?,
            i16::try_from(right)
                .map_err(|_error| invalid_durable_patch("small anchor right is out of range"))?,
            i16::try_from(bottom)
                .map_err(|_error| invalid_durable_patch("small anchor bottom is out of range"))?,
        )
        .map_err(Error::from),
        Some("full") => crate::Anchor::full(left, top, right, bottom).map_err(Error::from),
        _ => Err(invalid_durable_patch(
            "shape-anchor encoding must be small or full",
        )),
    }
}

fn apply_durable_structural(
    transaction: &mut Transaction,
    operation: &litchi_core::patch::PatchOperation,
    blobs: &litchi_core::patch::BlobBundle,
    artifact: &str,
) -> Result<()> {
    let expected_artifact = operation
        .preconditions
        .get("artifact_sha256")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| invalid_durable_patch("missing artifact hash precondition"))?;
    if expected_artifact != artifact {
        return Err(PackageError::InvalidFormat(
            "PPT durable structural patch source artifact does not match".into(),
        )
        .into());
    }
    let expected_order = operation
        .preconditions
        .get("order_sha256")
        .and_then(serde_json::Value::as_str)
        .ok_or_else(|| invalid_durable_patch("missing slide-order precondition"))?;
    if expected_order != transaction.order_fingerprint()? {
        return Err(PackageError::InvalidFormat(
            "PPT durable slide-order semantic precondition does not match".into(),
        )
        .into());
    }
    if operation.op.starts_with("document-structure.") {
        return apply_durable_document(transaction, operation, blobs);
    }
    let position = parse_target(&operation.target)?;
    match operation.op.as_str() {
        "slide-order.move" if operation.preconditions.len() == 2 => {
            let destination = operation
                .value
                .as_u64()
                .and_then(|value| usize::try_from(value).ok())
                .map(Position::new)
                .ok_or_else(|| invalid_durable_patch("slide destination must fit usize"))?;
            transaction.move_slide(position, destination)
        },
        "slide-order.remove" if operation.preconditions.len() == 3 && operation.value.is_null() => {
            transaction.require_position(position)?;
            let persist_id = transaction.document.slides()?[position.get()].persist_id();
            let record = persisted_record(&transaction.working, persist_id)?;
            let expected_slide = operation
                .preconditions
                .get("slide_sha256")
                .and_then(serde_json::Value::as_str)
                .ok_or_else(|| invalid_durable_patch("missing slide payload precondition"))?;
            if artifact_hash(&record) != expected_slide {
                return Err(PackageError::InvalidFormat(
                    "PPT durable slide payload precondition does not match".into(),
                )
                .into());
            }
            transaction.remove_slide(position)
        },
        "slide-order.insert" if operation.preconditions.len() == 3 => {
            let blob_hex = operation
                .value
                .as_str()
                .ok_or_else(|| invalid_durable_patch("slide insertion value must be a blob ID"))?;
            let blob = blob_by_hex(blobs, blob_hex)
                .ok_or_else(|| invalid_durable_patch("slide insertion blob is missing"))?;
            let payload = decode_payload(blob, transaction.source.limits)?;
            let expected_master = operation
                .preconditions
                .get("master_id")
                .and_then(serde_json::Value::as_u64)
                .and_then(|value| u32::try_from(value).ok())
                .ok_or_else(|| invalid_durable_patch("slide insertion master is invalid"))?;
            if payload.master_id != expected_master {
                return Err(invalid_durable_patch(
                    "slide insertion master precondition disagrees with its blob",
                ));
            }
            let plan = TransferPlan {
                target_artifact: artifact.to_string(),
                payload,
                reused_dependencies: Vec::new(),
                relationship_remaps: Vec::new(),
            };
            transaction.insert_transfer(position, &plan)
        },
        _ => Err(invalid_durable_patch(
            "unsupported structural operation vocabulary",
        )),
    }
}

fn blob_by_hex<'a>(blobs: &'a litchi_core::patch::BlobBundle, value: &str) -> Option<&'a [u8]> {
    blobs
        .ids()
        .find(|identifier| identifier.as_hex() == value)
        .and_then(|identifier| blobs.get(identifier))
}

fn parse_shape_target(value: &str) -> Result<crate::text_edit::Target> {
    let (slide, shape) = value
        .strip_prefix("slide:")
        .and_then(|suffix| suffix.split_once("/shape:"))
        .ok_or_else(|| invalid_durable_patch("invalid shape-text target"))?;
    let parse = |part: &str| {
        if part.is_empty()
            || !part.bytes().all(|byte| byte.is_ascii_digit())
            || (part != "0" && part.starts_with('0'))
        {
            return Err(invalid_durable_patch("invalid shape-text position"));
        }
        part.parse::<usize>()
            .map(Position::new)
            .map_err(|_error| invalid_durable_patch("shape-text position exceeds this platform"))
    };
    Ok(crate::text_edit::Target::new(parse(slide)?, parse(shape)?))
}

/// A `RoundTripCustomTableStyles12Atom` header the document validator
/// accepts (instance 0, no children, a consistent length).
fn require_custom_table_styles_record(record: &crate::Record) -> Result<()> {
    if record.record_type != crate::consts::RecordType::RoundTripCustomTableStyles12Atom
        || record.record_type_raw
            != crate::consts::RecordType::RoundTripCustomTableStyles12Atom.as_u16()
        || record.instance != 0
        || record.version > 0x000f
        || !record.children.is_empty()
        || usize::try_from(record.data_length).ok() != Some(record.data.len())
    {
        return Err(Error::Refused(Refusal::InvalidElementValue {
            element: "custom table styles",
            reason: "the record is not a RoundTripCustomTableStyles12Atom",
        }));
    }
    Ok(())
}

fn order_digest(transaction: &document_structure::Transaction) -> Result<[u8; 32]> {
    let slides = transaction.slides()?;
    Ok(digest_slide_ids(
        slides.iter().map(|slide| slide.slide_id()),
    ))
}

fn digest_slide_ids(ids: impl IntoIterator<Item = u32>) -> [u8; 32] {
    let bytes = ids
        .into_iter()
        .flat_map(u32::to_le_bytes)
        .collect::<Vec<_>>();
    let hex = litchi_core::patch::BlobId::of(&bytes).as_hex();
    let mut digest = [0_u8; 32];
    for (index, pair) in hex.as_bytes().as_chunks::<2>().0.iter().enumerate() {
        digest[index] = (hex_nibble(pair[0]) << 4) | hex_nibble(pair[1]);
    }
    digest
}

fn hex_digest(digest: [u8; 32]) -> String {
    use std::fmt::Write as _;
    let mut text = String::with_capacity(64);
    for byte in digest {
        let _result = write!(text, "{byte:02x}");
    }
    text
}

pub(crate) fn external_media_snapshot(
    snapshot: &Snapshot,
) -> Result<crate::external_media::Snapshot> {
    let limits = crate::external_media::Limits {
        max_root_bytes: snapshot
            .limits
            .max_record_bytes
            .min(snapshot.limits.max_input_bytes),
        max_record_bytes: snapshot.limits.max_record_bytes,
        max_depth: snapshot.limits.max_depth,
        max_records: snapshot.limits.max_records,
        max_owner_references: snapshot.limits.max_records,
    };
    crate::external_media::Snapshot::parse_with_limits(snapshot.document.bytes(), limits)
        .map_err(Error::from)
}

fn media_path_of(object: &crate::external_media::Object) -> Result<Option<String>> {
    match object {
        crate::external_media::Object::Movie(movie) => Ok(movie.video.path.clone()),
        crate::external_media::Object::LinkedAudio(audio) => Ok(audio.path.clone()),
        crate::external_media::Object::CdAudio(_)
        | crate::external_media::Object::EmbeddedWav(_) => Err(PackageError::InvalidFormat(
            "external media object does not own a linked path".into(),
        )
        .into()),
    }
}

fn publish_live_document(snapshot: &Snapshot, document: &[u8]) -> Result<Snapshot> {
    let mut editor = crate::embedded::object::Editor::open_records_arc_with_limit(
        snapshot.shared_bytes(),
        snapshot.limits.max_package_bytes,
        snapshot.limits,
    )?;
    let live = editor.persisted_record(snapshot.document_persist_id)?;
    if live.as_slice() != snapshot.document.bytes() {
        return Err(PackageError::Corrupted(
            "PPT live document changed before owner publication".into(),
        )
        .into());
    }
    editor.replace_persisted_record(snapshot.document_persist_id, document.to_vec())?;
    let bytes = editor.finish()?;
    crate::font::validate_unrelated_streams(snapshot.bytes(), &bytes)?;
    let published = Snapshot::from_bytes_with_options(bytes, snapshot.open_options())?;
    if published.document.slides() != snapshot.document.slides() {
        return Err(PackageError::Corrupted(
            "PPT live-document owner publication changed slide structure".into(),
        )
        .into());
    }
    Ok(published)
}

const fn hex_nibble(value: u8) -> u8 {
    match value {
        b'0'..=b'9' => value - b'0',
        b'a'..=b'f' => value - b'a' + 10,
        _ => 0,
    }
}

fn persisted_slides(
    bytes: &Arc<[u8]>,
    document: &document_structure::Snapshot,
    limits: SourceLimits,
) -> Result<BTreeMap<u32, Vec<u8>>> {
    let editor = crate::embedded::object::Editor::open_records_arc_with_limit(
        bytes.clone(),
        bytes.len().saturating_mul(4),
        limits,
    )?;
    editor_persisted_slides(&editor, document)
}

fn editor_persisted_slides(
    editor: &crate::embedded::object::Editor,
    document: &document_structure::Snapshot,
) -> Result<BTreeMap<u32, Vec<u8>>> {
    document
        .slides()
        .iter()
        .map(|slide| {
            editor
                .persisted_record(slide.persist_id())
                .map(|record| (slide.persist_id(), record))
                .map_err(Error::from)
        })
        .collect()
}

fn persisted_record(snapshot: &Snapshot, persist_id: u32) -> Result<Vec<u8>> {
    if snapshot.has_signature {
        return persisted_record_readonly(snapshot, persist_id);
    }
    crate::embedded::object::Editor::open_records_arc_with_limit(
        snapshot.shared_bytes(),
        snapshot.limits.max_package_bytes,
        snapshot.limits,
    )?
    .persisted_record(persist_id)
    .map_err(Error::from)
}

/// Reads one persisted record without opening the mutation-capable editor.
///
/// A signed CFB can be projected through the normal presentation parser, but
/// the append-only editor must reject it before it can publish a changed
/// package.  The read-only path therefore resolves the selected persist ID
/// from the presentation's retained mapping and slices the original record
/// bytes directly from the PowerPoint Document stream.  This preserves the
/// same source limits and leniency policy captured by the snapshot.
fn persisted_record_readonly(snapshot: &Snapshot, persist_id: u32) -> Result<Vec<u8>> {
    let mut package =
        Reader::from_reader_with_options(Cursor::new(snapshot.bytes()), snapshot.open_options())?;
    let presentation = package.presentation_with_options(snapshot.open_options())?;
    let factory = crate::slide::SlideFactory::new_shared_with_limits(
        &presentation.inner.powerpoint_document,
        &presentation.inner.persist_mapping,
        &presentation.inner.slide_directory,
        presentation.source_limits(),
    );
    let slide = factory.parse_slide(persist_id)?;
    let document = slide.doc_data();
    let tail = document.get(slide.offset..).ok_or_else(|| {
        PackageError::Corrupted(
            "PPT selected persist record offset exceeds its document stream".into(),
        )
    })?;
    let (_, consumed) = crate::Record::parse(tail, 0, &presentation.source_limits())?;
    let end = slide.offset.checked_add(consumed).ok_or_else(|| {
        PackageError::Corrupted("PPT selected persist record length overflows".into())
    })?;
    document
        .get(slide.offset..end)
        .map(ToOwned::to_owned)
        .ok_or_else(|| {
            PackageError::Corrupted(
                "PPT selected persist record exceeds its document stream".into(),
            )
            .into()
        })
}

/// Reads one persisted record of any kind, such as a master, by persist ID.
///
/// An unsigned package reads through the append-only editor, as
/// [`persisted_record`] does; a signed one slices the record from the
/// PowerPoint Document stream at its retained persist offset.
fn persisted_object_record(snapshot: &Snapshot, persist_id: u32) -> Result<Vec<u8>> {
    if !snapshot.has_signature {
        return persisted_record(snapshot, persist_id);
    }
    let mut package =
        Reader::from_reader_with_options(Cursor::new(snapshot.bytes()), snapshot.open_options())?;
    let presentation = package.presentation_with_options(snapshot.open_options())?;
    let offset = presentation
        .inner
        .persist_mapping
        .get_offset(persist_id)
        .and_then(|offset| usize::try_from(offset).ok())
        .ok_or_else(|| {
            PackageError::Corrupted("PPT selected persist record has no directory entry".into())
        })?;
    let document = presentation.document_stream();
    let tail = document.get(offset..).ok_or_else(|| {
        PackageError::Corrupted(
            "PPT selected persist record offset exceeds its document stream".into(),
        )
    })?;
    let (_, consumed) = crate::Record::parse(tail, 0, &presentation.source_limits())?;
    let end = offset.checked_add(consumed).ok_or_else(|| {
        PackageError::Corrupted("PPT selected persist record length overflows".into())
    })?;
    document
        .get(offset..end)
        .map(ToOwned::to_owned)
        .ok_or_else(|| {
            PackageError::Corrupted(
                "PPT selected persist record exceeds its document stream".into(),
            )
            .into()
        })
}

fn inspect_slide_hidden(snapshot: &Snapshot, position: Position) -> Result<bool> {
    let slide = snapshot
        .document
        .slides()
        .get(position.get())
        .copied()
        .ok_or(Error::Refused(Refusal::SlideNotFound { position }))?;
    let bytes = persisted_record(snapshot, slide.persist_id())?;
    let (root, consumed) = crate::Record::parse(&bytes, 0, &snapshot.limits)?;
    if consumed != bytes.len() || root.record_type != crate::RecordType::Slide {
        return Err(PackageError::Corrupted(
            "PPT slide visibility owner is not one complete SlideContainer".into(),
        )
        .into());
    }
    match optional_slide_hidden_atom(&root, position)? {
        Some(atom) => parse_slide_hidden_atom(atom, position),
        // A slide without slide-show information is shown ([MS-PPT] 2.5.1:
        // the atom is optional; `fHidden` defaults to clear).
        None => Ok(false),
    }
}

/// The slide's unique `SlideShowSlideInfoAtom`, or `None` when it has none;
/// several are a refusal.
fn optional_slide_hidden_atom(
    root: &crate::Record,
    position: Position,
) -> Result<Option<&crate::Record>> {
    let mut atoms = root
        .children
        .iter()
        .filter(|child| child.record_type == crate::RecordType::SSSlideInfoAtom);
    let atom = atoms.next();
    if atoms.next().is_some() {
        return Err(Error::Refused(Refusal::UnsupportedSlideVisibility {
            position,
        }));
    }
    Ok(atom)
}

/// A `SlideShowSlideInfoAtom` for a slide that had none: no transition, no
/// sound, manual advance and medium speed, what the PPT writer emits for
/// default slide timing. The caller sets `fHidden`.
fn new_slide_hidden_atom(limits: SourceLimits) -> Result<crate::Record> {
    let timing = crate::writer::slide_timing::SlideTiming::default();
    let bytes = crate::writer::slide_timing::build_slide_timing(&timing).map_err(|error| {
        PackageError::Corrupted(format!("PPT slide-show information atom: {error}"))
    })?;
    let (atom, consumed) = crate::Record::parse(&bytes, 0, &limits)?;
    if consumed != bytes.len() {
        return Err(PackageError::Corrupted(
            "PPT slide-show information atom is not one complete record".into(),
        )
        .into());
    }
    Ok(atom)
}

fn parse_slide_hidden_atom(atom: &crate::Record, position: Position) -> Result<bool> {
    const RESERVED_FLAGS: u16 =
        (1 << 1) | (1 << 3) | (1 << 5) | (1 << 7) | (1 << 9) | (1 << 11) | (0b111 << 13);

    if atom.version != 0 || atom.instance != 0 || !atom.children.is_empty() || atom.data.len() != 16
    {
        return Err(Error::Refused(Refusal::UnsupportedSlideVisibility {
            position,
        }));
    }
    let flags = u16::from_le_bytes([atom.data[10], atom.data[11]]);
    if flags & RESERVED_FLAGS != 0 {
        return Err(Error::Refused(Refusal::UnsupportedSlideVisibility {
            position,
        }));
    }
    Ok(flags & (1 << 2) != 0)
}

fn replace_slide_hidden(
    snapshot: &Snapshot,
    position: Position,
    replacement: bool,
) -> Result<Snapshot> {
    let slide = snapshot
        .document
        .slides()
        .get(position.get())
        .copied()
        .ok_or(Error::Refused(Refusal::SlideNotFound { position }))?;
    let source_record = persisted_record(snapshot, slide.persist_id())?;
    let (mut root, consumed) = crate::Record::parse(&source_record, 0, &snapshot.limits)?;
    if consumed != source_record.len() || root.record_type != crate::RecordType::Slide {
        return Err(PackageError::Corrupted(
            "PPT slide visibility owner is not one complete SlideContainer".into(),
        )
        .into());
    }
    if encode_record(&root)? != source_record {
        return Err(PackageError::Corrupted(
            "PPT slide visibility source record is not byte-exactly re-encodable".into(),
        )
        .into());
    }
    let inserted = match optional_slide_hidden_atom(&root, position)? {
        Some(atom) => {
            let _current = parse_slide_hidden_atom(atom, position)?;
            false
        },
        None => {
            if !replacement {
                // Absent slide-show information already reads as shown.
                return Ok(snapshot.clone());
            }
            // [MS-PPT] 2.5.1 places the optional atom directly after the
            // SlideAtom, which must be the container's first child.
            if root
                .children
                .first()
                .is_none_or(|child| child.record_type != crate::RecordType::SlideAtom)
            {
                return Err(Error::Refused(Refusal::UnsupportedSlideVisibility {
                    position,
                }));
            }
            let atom = new_slide_hidden_atom(snapshot.limits)?;
            root.children.insert(1, atom);
            true
        },
    };
    let atom_index = root
        .children
        .iter()
        .position(|child| child.record_type == crate::RecordType::SSSlideInfoAtom)
        .ok_or(Error::Refused(Refusal::UnsupportedSlideVisibility {
            position,
        }))?;
    let atom = &mut root.children[atom_index];
    let before = atom.data.clone();
    let mut flags = u16::from_le_bytes([atom.data[10], atom.data[11]]);
    if replacement {
        flags |= 1 << 2;
    } else {
        flags &= !(1 << 2);
    }
    atom.data[10..12].copy_from_slice(&flags.to_le_bytes());
    if atom.data[..10] != before[..10] || atom.data[12..] != before[12..] {
        return Err(PackageError::Corrupted(
            "PPT slide visibility rewrite changed non-visibility transition fields".into(),
        )
        .into());
    }
    let rewritten_record = encode_record(&root)?;
    // An in-place flag change keeps the record length; an inserted atom adds
    // exactly its 8-byte header and 16-byte payload.
    let expected_len = source_record.len() + if inserted { 24 } else { 0 };
    if rewritten_record.len() != expected_len {
        return Err(PackageError::Corrupted(
            "PPT slide visibility rewrite changed record length".into(),
        )
        .into());
    }
    let mut editor = crate::embedded::object::Editor::open_records_arc_with_limit(
        snapshot.shared_bytes(),
        snapshot.limits.max_package_bytes,
        snapshot.limits,
    )?;
    editor.replace_persisted_record(slide.persist_id(), rewritten_record)?;
    let bytes = editor.finish()?;
    crate::font::validate_unrelated_streams(snapshot.bytes(), &bytes)?;
    let published = Snapshot::from_bytes_with_options(bytes, snapshot.open_options())?;
    if published.document != snapshot.document || published.slide_hidden(position)? != replacement {
        return Err(PackageError::Corrupted(
            "PPT slide visibility replacement did not round-trip exactly".into(),
        )
        .into());
    }
    Ok(published)
}

fn inspect_slide_advance(snapshot: &Snapshot, position: Position) -> Result<SlideAdvance> {
    let slide = snapshot
        .document
        .slides()
        .get(position.get())
        .copied()
        .ok_or(Error::Refused(Refusal::SlideNotFound { position }))?;
    let bytes = persisted_record(snapshot, slide.persist_id())?;
    let (root, consumed) = crate::Record::parse(&bytes, 0, &snapshot.limits)?;
    if consumed != bytes.len() || root.record_type != crate::RecordType::Slide {
        return Err(PackageError::Corrupted(
            "PPT slide advance owner is not one complete SlideContainer".into(),
        )
        .into());
    }
    let atom = unique_slide_advance_atom(&root, position)?;
    parse_slide_advance_atom(atom, position)
}

fn unique_slide_advance_atom(root: &crate::Record, position: Position) -> Result<&crate::Record> {
    let mut atoms = root
        .children
        .iter()
        .filter(|child| child.record_type == crate::RecordType::SSSlideInfoAtom);
    let atom = atoms
        .next()
        .ok_or(Error::Refused(Refusal::UnsupportedSlideAdvance {
            position,
        }))?;
    if atoms.next().is_some() {
        return Err(Error::Refused(Refusal::UnsupportedSlideAdvance {
            position,
        }));
    }
    Ok(atom)
}

fn parse_slide_advance_atom(atom: &crate::Record, position: Position) -> Result<SlideAdvance> {
    const RESERVED_FLAGS: u16 =
        (1 << 1) | (1 << 3) | (1 << 5) | (1 << 7) | (1 << 9) | (1 << 11) | (0b111 << 13);

    if atom.version != 0 || atom.instance != 0 || !atom.children.is_empty() || atom.data.len() != 16
    {
        return Err(Error::Refused(Refusal::UnsupportedSlideAdvance {
            position,
        }));
    }
    let flags = u16::from_le_bytes([atom.data[10], atom.data[11]]);
    let delay_ms = read_payload_u32(&atom.data, 0)?;
    if flags & RESERVED_FLAGS != 0 || delay_ms > MAX_SLIDE_ADVANCE_MS {
        return Err(Error::Refused(Refusal::UnsupportedSlideAdvance {
            position,
        }));
    }
    SlideAdvance::new(flags & 1 != 0, flags & (1 << 10) != 0, delay_ms)
}

fn replace_slide_advance(
    snapshot: &Snapshot,
    position: Position,
    replacement: SlideAdvance,
) -> Result<Snapshot> {
    let slide = snapshot
        .document
        .slides()
        .get(position.get())
        .copied()
        .ok_or(Error::Refused(Refusal::SlideNotFound { position }))?;
    let source_record = persisted_record(snapshot, slide.persist_id())?;
    let (mut root, consumed) = crate::Record::parse(&source_record, 0, &snapshot.limits)?;
    if consumed != source_record.len() || root.record_type != crate::RecordType::Slide {
        return Err(PackageError::Corrupted(
            "PPT slide advance owner is not one complete SlideContainer".into(),
        )
        .into());
    }
    if encode_record(&root)? != source_record {
        return Err(PackageError::Corrupted(
            "PPT slide advance source record is not byte-exactly re-encodable".into(),
        )
        .into());
    }
    let atom_index = {
        let _current =
            parse_slide_advance_atom(unique_slide_advance_atom(&root, position)?, position)?;
        root.children
            .iter()
            .position(|child| child.record_type == crate::RecordType::SSSlideInfoAtom)
            .ok_or(Error::Refused(Refusal::UnsupportedSlideAdvance {
                position,
            }))?
    };
    let atom = &mut root.children[atom_index];
    let before = atom.data.clone();
    atom.data[0..4].copy_from_slice(&replacement.delay_ms.to_le_bytes());
    let mut flags = u16::from_le_bytes([atom.data[10], atom.data[11]]);
    flags &= !((1 << 0) | (1 << 10));
    flags |= u16::from(replacement.manual);
    flags |= u16::from(replacement.automatic) << 10;
    atom.data[10..12].copy_from_slice(&flags.to_le_bytes());
    if atom.data[4..10] != before[4..10] || atom.data[12..16] != before[12..16] {
        return Err(PackageError::Corrupted(
            "PPT slide advance rewrite changed preserved transition fields".into(),
        )
        .into());
    }
    let rewritten_record = encode_record(&root)?;
    if rewritten_record.len() != source_record.len() {
        return Err(PackageError::Corrupted(
            "PPT fixed-width slide advance rewrite changed record length".into(),
        )
        .into());
    }
    let mut editor = crate::embedded::object::Editor::open_records_arc_with_limit(
        snapshot.shared_bytes(),
        snapshot.limits.max_package_bytes,
        snapshot.limits,
    )?;
    editor.replace_persisted_record(slide.persist_id(), rewritten_record)?;
    let bytes = editor.finish()?;
    crate::font::validate_unrelated_streams(snapshot.bytes(), &bytes)?;
    let published = Snapshot::from_bytes_with_options(bytes, snapshot.open_options())?;
    if published.document != snapshot.document || published.slide_advance(position)? != replacement
    {
        return Err(PackageError::Corrupted(
            "PPT slide advance replacement did not round-trip exactly".into(),
        )
        .into());
    }
    Ok(published)
}

fn slide_transition_visual_from_bytes(bytes: [u8; 3]) -> Option<SlideTransitionVisual> {
    let (transition_type, direction, speed) = crate::transition::decode_visual(bytes)?;
    Some(SlideTransitionVisual {
        transition_type,
        direction,
        speed,
        wire: bytes,
    })
}

fn inspect_slide_transition_visual(
    snapshot: &Snapshot,
    position: Position,
) -> Result<SlideTransitionVisual> {
    let slide = snapshot
        .document
        .slides()
        .get(position.get())
        .copied()
        .ok_or(Error::Refused(Refusal::SlideNotFound { position }))?;
    let bytes = persisted_record(snapshot, slide.persist_id())?;
    let (root, consumed) = crate::Record::parse(&bytes, 0, &snapshot.limits)?;
    if consumed != bytes.len() || root.record_type != crate::RecordType::Slide {
        return Err(PackageError::Corrupted(
            "PPT visual transition owner is not one complete SlideContainer".into(),
        )
        .into());
    }
    parse_slide_transition_visual_atom(
        unique_slide_transition_visual_atom(&root, position)?,
        position,
    )
}

fn unique_slide_transition_visual_atom(
    root: &crate::Record,
    position: Position,
) -> Result<&crate::Record> {
    let mut atoms = root
        .children
        .iter()
        .filter(|child| child.record_type == crate::RecordType::SSSlideInfoAtom);
    let atom = atoms
        .next()
        .ok_or(Error::Refused(Refusal::UnsupportedSlideTransitionVisual {
            position,
        }))?;
    if atoms.next().is_some() {
        return Err(Error::Refused(Refusal::UnsupportedSlideTransitionVisual {
            position,
        }));
    }
    Ok(atom)
}

fn parse_slide_transition_visual_atom(
    atom: &crate::Record,
    position: Position,
) -> Result<SlideTransitionVisual> {
    if atom.version != 0 || atom.instance != 0 || !atom.children.is_empty() || atom.data.len() != 16
    {
        return Err(Error::Refused(Refusal::UnsupportedSlideTransitionVisual {
            position,
        }));
    }
    slide_transition_visual_from_bytes([atom.data[8], atom.data[9], atom.data[12]]).ok_or(
        Error::Refused(Refusal::UnsupportedSlideTransitionVisual { position }),
    )
}

fn replace_slide_transition_visual(
    snapshot: &Snapshot,
    position: Position,
    replacement: SlideTransitionVisual,
) -> Result<Snapshot> {
    let slide = snapshot
        .document
        .slides()
        .get(position.get())
        .copied()
        .ok_or(Error::Refused(Refusal::SlideNotFound { position }))?;
    let source_record = persisted_record(snapshot, slide.persist_id())?;
    let (mut root, consumed) = crate::Record::parse(&source_record, 0, &snapshot.limits)?;
    if consumed != source_record.len() || root.record_type != crate::RecordType::Slide {
        return Err(PackageError::Corrupted(
            "PPT visual transition owner is not one complete SlideContainer".into(),
        )
        .into());
    }
    if encode_record(&root)? != source_record {
        return Err(PackageError::Corrupted(
            "PPT visual transition source record is not byte-exactly re-encodable".into(),
        )
        .into());
    }
    let atom_index = {
        let _current = parse_slide_transition_visual_atom(
            unique_slide_transition_visual_atom(&root, position)?,
            position,
        )?;
        root.children
            .iter()
            .position(|child| child.record_type == crate::RecordType::SSSlideInfoAtom)
            .ok_or(Error::Refused(Refusal::UnsupportedSlideTransitionVisual {
                position,
            }))?
    };
    let atom = &mut root.children[atom_index];
    let before = atom.data.clone();
    let [effect_direction, effect_type, speed] = transition_visual_bytes(replacement);
    atom.data[8] = effect_direction;
    atom.data[9] = effect_type;
    atom.data[12] = speed;
    if atom.data[..8] != before[..8]
        || atom.data[10..12] != before[10..12]
        || atom.data[13..16] != before[13..16]
    {
        return Err(PackageError::Corrupted(
            "PPT visual transition rewrite changed timing, sound, flags, or unused fields".into(),
        )
        .into());
    }
    let rewritten_record = encode_record(&root)?;
    if rewritten_record.len() != source_record.len() {
        return Err(PackageError::Corrupted(
            "PPT fixed-width visual transition rewrite changed record length".into(),
        )
        .into());
    }
    let mut editor = crate::embedded::object::Editor::open_records_arc_with_limit(
        snapshot.shared_bytes(),
        snapshot.limits.max_package_bytes,
        snapshot.limits,
    )?;
    editor.replace_persisted_record(slide.persist_id(), rewritten_record)?;
    let bytes = editor.finish()?;
    crate::font::validate_unrelated_streams(snapshot.bytes(), &bytes)?;
    let published = Snapshot::from_bytes_with_options(bytes, snapshot.open_options())?;
    if published.document != snapshot.document
        || published.slide_transition_visual(position)? != replacement
    {
        return Err(PackageError::Corrupted(
            "PPT visual transition replacement did not round-trip exactly".into(),
        )
        .into());
    }
    Ok(published)
}

fn expected_persisted_slides(
    slides: &[document_structure::Slide],
    existing: &BTreeMap<u32, Vec<u8>>,
    inserted: &BTreeMap<u32, Vec<u8>>,
) -> Result<BTreeMap<u32, Vec<u8>>> {
    slides
        .iter()
        .map(|slide| {
            existing
                .get(&slide.persist_id())
                .or_else(|| inserted.get(&slide.persist_id()))
                .cloned()
                .map(|record| (slide.persist_id(), record))
                .ok_or_else(|| {
                    PackageError::Corrupted(
                        "published slide list references an unstaged persist record".into(),
                    )
                    .into()
                })
        })
        .collect()
}

fn require_master(document: &document_structure::Transaction, master_id: u32) -> Result<()> {
    if document
        .masters()?
        .iter()
        .any(|master| master.master_id() == master_id)
    {
        Ok(())
    } else {
        Err(Error::Refused(Refusal::MissingMasterDependency {
            master_id,
        }))
    }
}

fn require_master_closure(target: &Snapshot, donor: &Snapshot, master_id: u32) -> Result<()> {
    let target_master = target
        .document
        .masters()
        .iter()
        .find(|master| master.master_id() == master_id)
        .copied()
        .ok_or(Error::Refused(Refusal::MissingMasterDependency {
            master_id,
        }))?;
    let donor_master = donor
        .document
        .masters()
        .iter()
        .find(|master| master.master_id() == master_id)
        .copied()
        .ok_or_else(|| {
            PackageError::Corrupted(
                "PPT transferred slide references a donor master that is absent".into(),
            )
        })?;
    if persisted_record(target, target_master.persist_id())?
        == persisted_record(donor, donor_master.persist_id())?
    {
        Ok(())
    } else {
        Err(Error::Refused(Refusal::MismatchedMasterDependency {
            master_id,
        }))
    }
}

fn next_persist_id(source: &Snapshot, inserted: &BTreeMap<u32, Vec<u8>>) -> Result<u32> {
    let editor = crate::embedded::object::Editor::open_records_arc_with_limit(
        source.shared_bytes(),
        source.limits.max_package_bytes,
        source.limits,
    )?;
    editor
        .persist_ids()
        .into_iter()
        .chain(inserted.keys().copied())
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .filter(|identifier| *identifier <= 0x000f_ffff)
        .ok_or_else(|| {
            PackageError::ResourceLimit("PPT persist ID space is exhausted".into()).into()
        })
}

fn next_slide_id(document: &document_structure::Transaction) -> Result<u32> {
    document
        .slides()?
        .iter()
        .map(|slide| slide.slide_id())
        .max()
        .unwrap_or(0xff)
        .checked_add(1)
        .filter(|identifier| (0x100..=0x7fff_ffff).contains(identifier))
        .ok_or_else(|| PackageError::ResourceLimit("PPT slide ID space is exhausted".into()).into())
}

fn rewrite_slide_persist(
    group: &mut [crate::records::Record],
    persist_id: u32,
    slide_id: u32,
) -> Result<()> {
    let atom = group
        .first_mut()
        .ok_or_else(|| PackageError::Corrupted("transferred slide group is empty".into()))?;
    if atom.record_type != crate::RecordType::SlidePersistAtom || atom.data.len() != 20 {
        return Err(PackageError::Corrupted(
            "transferred slide group has an invalid SlidePersistAtom".into(),
        )
        .into());
    }
    atom.data[0..4].copy_from_slice(&persist_id.to_le_bytes());
    atom.data[12..16].copy_from_slice(&slide_id.to_le_bytes());
    Ok(())
}

fn slide_payload(group: Vec<crate::records::Record>, record: Vec<u8>) -> Result<SlidePayload> {
    if group.is_empty()
        || group[0].record_type != crate::RecordType::SlidePersistAtom
        || group[0].data.len() != 20
    {
        return Err(PackageError::Corrupted(
            "slide transfer source has an invalid slide-list group".into(),
        )
        .into());
    }
    let (root, consumed) = crate::Record::parse(&record, 0, &SourceLimits::default())?;
    if consumed != record.len() || root.record_type != crate::RecordType::Slide {
        return Err(PackageError::Corrupted(
            "slide transfer persist record is not one SlideContainer".into(),
        )
        .into());
    }
    let mut atoms = root
        .children
        .iter()
        .filter(|child| child.record_type == crate::RecordType::SlideAtom);
    let atom = atoms
        .next()
        .ok_or_else(|| PackageError::Corrupted("slide transfer source has no SlideAtom".into()))?;
    if atoms.next().is_some() || atom.data.len() != 24 {
        return Err(PackageError::Corrupted(
            "slide transfer source has an ambiguous SlideAtom".into(),
        )
        .into());
    }
    let master_id = read_payload_u32(&atom.data, 12)?;
    let notes_id = read_payload_u32(&atom.data, 16)?;
    let _ = notes_id;
    Ok(SlidePayload {
        group,
        record,
        master_id,
        pictures: Vec::new(),
    })
}

struct TransferClosure {
    dependencies: Vec<TransferDependency>,
    remaps: Vec<RelationshipRemap>,
    pictures: Vec<TransferredPicture>,
}

fn require_portable_slide(
    target: &Snapshot,
    donor: &Snapshot,
    record: &mut Vec<u8>,
) -> Result<TransferClosure> {
    let source_limits = target.limits();
    let (mut root, consumed) = crate::Record::parse(record, 0, &source_limits)?;
    if consumed != record.len() || root.record_type != crate::RecordType::Slide {
        return Err(PackageError::Corrupted(
            "slide transfer persist record is not one SlideContainer".into(),
        )
        .into());
    }
    let atom = root
        .children
        .iter()
        .find(|child| child.record_type == crate::RecordType::SlideAtom)
        .ok_or_else(|| PackageError::Corrupted("slide transfer source has no SlideAtom".into()))?;
    if read_payload_u32(&atom.data, 16)? != 0 {
        return Err(Error::Refused(Refusal::UnsupportedSlideDependency {
            dependency: TransferDependency::SpeakerNotes,
        }));
    }
    let mut reused = Vec::new();
    let comments = crate::comments::parse_slide_comments(&root, &source_limits)?;
    if !comments.is_empty() {
        let target_authors =
            crate::comments::Authors::parse(target.document.record(), &source_limits)?;
        if comments.iter().any(|comment| {
            target_authors
                .find(&comment.author)
                .and_then(|author| author.comment_index_seed)
                .is_none_or(|seed| comment.index > seed)
        }) {
            return Err(Error::Refused(Refusal::UnsupportedSlideDependency {
                dependency: TransferDependency::CommentAuthorCatalog,
            }));
        }
        reused.push(TransferDependency::CommentAuthorCatalog);
    }

    let relationships = scan_slide_relationships(&root)?;
    if relationships.active_action {
        return Err(Error::Refused(Refusal::UnsupportedSlideDependency {
            dependency: TransferDependency::ActiveOrUnknownExternalObject,
        }));
    }
    let hyperlink_remaps = hyperlink_remaps(target, donor, &relationships.hyperlink_ids)?;
    if !relationships.hyperlink_ids.is_empty() {
        reused.push(TransferDependency::HyperlinkAction);
    }
    let sound_remaps = sound_remaps(target, donor, &relationships.sound_ids)?;
    if !relationships.sound_ids.is_empty() {
        reused.push(TransferDependency::ExternalObjectRelationship);
    }
    let has_drawing = contains_record_type(&root, crate::RecordType::PPDrawing);
    let (picture_remaps, pictures) = picture_remaps(target, donor, &relationships.blip_ids)?;
    if !relationships.blip_ids.is_empty() {
        reused.push(TransferDependency::PictureStore);
    }
    let drawing_remaps = drawing_remaps(target, &relationships)?;
    let has_external = !relationships.external_object_ids.is_empty();
    let external_remaps =
        external_object_remaps(target, donor, &relationships.external_object_ids)?;
    if has_external {
        if external_owner_contains_active_or_unknown(donor.document.record()) {
            return Err(Error::Refused(Refusal::UnsupportedSlideDependency {
                dependency: TransferDependency::ActiveOrUnknownExternalObject,
            }));
        }
        reused.push(TransferDependency::ExternalObjectRelationship);
    }
    rewrite_slide_relationships(
        &mut root,
        &hyperlink_remaps,
        &sound_remaps,
        &external_remaps,
        &picture_remaps,
        &drawing_remaps,
    )?;
    *record = encode_record(&root)?;
    if has_drawing {
        reused.push(TransferDependency::DrawingGroup);
    }
    reused.sort_by_key(|dependency| transfer_dependency_order(*dependency));
    reused.dedup();
    let mut remaps = relationship_remaps(
        &hyperlink_remaps,
        &sound_remaps,
        &external_remaps,
        &picture_remaps,
        &drawing_remaps,
    );
    remaps.sort_by_key(|remap| {
        (
            transfer_dependency_order(remap.dependency),
            remap.source_id,
            remap.target_id,
        )
    });
    Ok(TransferClosure {
        dependencies: reused,
        remaps,
        pictures,
    })
}

fn relationship_remaps(
    hyperlinks: &BTreeMap<u32, u32>,
    sounds: &BTreeMap<u32, u32>,
    external_objects: &BTreeMap<u32, u32>,
    pictures: &BTreeMap<u32, u32>,
    drawings: &BTreeMap<u32, u32>,
) -> Vec<RelationshipRemap> {
    let mut output = Vec::new();
    output.extend(hyperlinks.iter().filter_map(|(source_id, target_id)| {
        (source_id != target_id).then_some(RelationshipRemap {
            dependency: TransferDependency::HyperlinkAction,
            source_id: *source_id,
            target_id: *target_id,
        })
    }));
    output.extend(sounds.iter().filter_map(|(source_id, target_id)| {
        (source_id != target_id).then_some(RelationshipRemap {
            dependency: TransferDependency::ExternalObjectRelationship,
            source_id: *source_id,
            target_id: *target_id,
        })
    }));
    output.extend(
        external_objects
            .iter()
            .filter_map(|(source_id, target_id)| {
                (source_id != target_id).then_some(RelationshipRemap {
                    dependency: TransferDependency::ExternalObjectRelationship,
                    source_id: *source_id,
                    target_id: *target_id,
                })
            }),
    );
    output.extend(pictures.iter().filter_map(|(source_id, target_id)| {
        (source_id != target_id).then_some(RelationshipRemap {
            dependency: TransferDependency::PictureStore,
            source_id: *source_id,
            target_id: *target_id,
        })
    }));
    output.extend(drawings.iter().filter_map(|(source_id, target_id)| {
        (source_id != target_id).then_some(RelationshipRemap {
            dependency: TransferDependency::DrawingGroup,
            source_id: *source_id,
            target_id: *target_id,
        })
    }));
    output
}

fn read_payload_u32(bytes: &[u8], offset: usize) -> Result<u32> {
    bytes
        .get(offset..offset.saturating_add(4))
        .and_then(|value| <[u8; 4]>::try_from(value).ok())
        .map(u32::from_le_bytes)
        .ok_or_else(|| PackageError::Corrupted("slide dependency field is truncated".into()).into())
}

fn contains_record_type(record: &crate::Record, target: crate::RecordType) -> bool {
    record.record_type == target
        || record
            .children
            .iter()
            .any(|child| contains_record_type(child, target))
}

#[derive(Default)]
struct SlideRelationships {
    hyperlink_ids: BTreeSet<u32>,
    sound_ids: BTreeSet<u32>,
    external_object_ids: BTreeSet<u32>,
    blip_ids: BTreeSet<u32>,
    shape_ids: BTreeSet<u32>,
    connector_shape_ids: BTreeSet<u32>,
    other_shape_reference_ids: BTreeSet<u32>,
    active_action: bool,
    unsupported_dependencies: BTreeSet<TransferDependency>,
}

fn scan_slide_relationships(root: &crate::Record) -> Result<SlideRelationships> {
    let mut relationships = SlideRelationships::default();
    scan_ppt_record(root, &mut relationships)?;
    Ok(relationships)
}

fn scan_ppt_record(record: &crate::Record, output: &mut SlideRelationships) -> Result<()> {
    if record.record_type == crate::RecordType::InteractiveInfoAtom {
        scan_interactive_atom(&record.data, output)?;
    } else if record.record_type == crate::RecordType::ExternalObjectRefAtom {
        scan_external_reference(&record.data, output)?;
    } else if record.record_type == crate::RecordType::PPDrawing {
        scan_officeart_records(&record.data, output)?;
    } else if record.record_type == crate::RecordType::BinaryTagData {
        scan_host_records(&record.data, output)?;
    } else if matches!(
        record.record_type,
        crate::RecordType::AnimationInfo
            | crate::RecordType::BuildList
            | crate::RecordType::ChartBuild
            | crate::RecordType::DiagramBuild
            | crate::RecordType::ParaBuild
    ) {
        output
            .unsupported_dependencies
            .insert(TransferDependency::AnimationBuildGraph);
    }
    for child in &record.children {
        scan_ppt_record(child, output)?;
    }
    Ok(())
}

fn option_blip_references(instance: u16, payload: &[u8]) -> Result<BTreeSet<u32>> {
    let property_bytes = usize::from(instance)
        .checked_mul(6)
        .ok_or_else(|| PackageError::Corrupted("OfficeArt property count overflow".into()))?;
    let properties = payload
        .get(..property_bytes)
        .ok_or_else(|| PackageError::Corrupted("OfficeArt property table is truncated".into()))?;
    let mut references = BTreeSet::new();
    for property in properties.as_chunks::<6>().0.iter() {
        if u16::from_le_bytes([property[0], property[1]]) & 0x4000 == 0 {
            continue;
        }
        let blip_id = u32::from_le_bytes([property[2], property[3], property[4], property[5]]);
        if blip_id != 0 {
            references.insert(blip_id);
        }
    }
    Ok(references)
}

fn scan_linked_shape_references(
    instance: u16,
    payload: &[u8],
    output: &mut SlideRelationships,
) -> Result<()> {
    let property_bytes = usize::from(instance)
        .checked_mul(6)
        .ok_or_else(|| PackageError::Corrupted("OfficeArt property count overflow".into()))?;
    let properties = payload
        .get(..property_bytes)
        .ok_or_else(|| PackageError::Corrupted("OfficeArt property table is truncated".into()))?;
    for property in properties.as_chunks::<6>().0.iter() {
        let opid = u16::from_le_bytes([property[0], property[1]]) & 0x3fff;
        if opid == 0x008a {
            let shape_id = u32::from_le_bytes([property[2], property[3], property[4], property[5]]);
            if shape_id != 0 {
                output.other_shape_reference_ids.insert(shape_id);
            }
        }
    }
    Ok(())
}

fn scan_connector_rule(payload: &[u8], output: &mut SlideRelationships) -> Result<()> {
    if payload.len() != 24 {
        return Err(PackageError::Corrupted(
            "OfficeArtFConnectorRule payload must contain 24 bytes".into(),
        )
        .into());
    }
    for offset in [4, 8, 12] {
        output
            .connector_shape_ids
            .insert(read_payload_u32(payload, offset)?);
    }
    Ok(())
}

fn scan_single_shape_rule(
    payload: &[u8],
    context: &str,
    output: &mut SlideRelationships,
) -> Result<()> {
    if payload.len() != 8 {
        return Err(
            PackageError::Corrupted(format!("{context} payload must contain eight bytes")).into(),
        );
    }
    let shape_id = read_payload_u32(payload, 4)?;
    if shape_id != 0 {
        output.other_shape_reference_ids.insert(shape_id);
    }
    Ok(())
}

fn scan_deleted_shape_rule(payload: &[u8], output: &mut SlideRelationships) -> Result<()> {
    if payload.len() != 4 {
        return Err(PackageError::Corrupted(
            "OfficeArtFPSPL payload must contain four bytes".into(),
        )
        .into());
    }
    let shape_id = read_payload_u32(payload, 0)? & 0x3fff_ffff;
    if shape_id != 0 {
        output.other_shape_reference_ids.insert(shape_id);
    }
    Ok(())
}

fn scan_drawing_selection(payload: &[u8], output: &mut SlideRelationships) -> Result<()> {
    if payload.len() < 12 || !(payload.len() - 12).is_multiple_of(4) {
        return Err(PackageError::Corrupted(
            "OfficeArtFDGSL payload has an invalid shape-list width".into(),
        )
        .into());
    }
    for offset in (8..payload.len()).step_by(4) {
        let shape_id = read_payload_u32(payload, offset)?;
        if shape_id != 0 {
            output.other_shape_reference_ids.insert(shape_id);
        }
    }
    Ok(())
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct PictureResource {
    kind: litchi_odraw::image::Kind,
    data: Vec<u8>,
    blip: Vec<u8>,
}

struct PictureInventory {
    slots: u32,
    resources: BTreeMap<u32, PictureResource>,
}

fn picture_inventory(snapshot: &Snapshot) -> Result<PictureInventory> {
    let mut package =
        Reader::from_reader_with_options(Cursor::new(snapshot.bytes()), snapshot.open_options())?;
    let presentation = package.presentation_with_options(snapshot.open_options())?;
    let slots = presentation
        .image_store()?
        .map_or(0, |store| u32::from(store.len()));
    let mut resources = BTreeMap::new();
    for image in presentation.images()? {
        let id = image
            .index()
            .checked_add(1)
            .and_then(|index| u32::try_from(index).ok())
            .ok_or_else(|| PackageError::Corrupted("PPT BLIP index exceeds u32".into()))?;
        let data = image.data().map_err(PackageError::from)?.to_vec();
        let mut blip = Vec::new();
        litchi_odraw::image::write::copy(
            &mut blip,
            &image.blip().map_err(PackageError::from)?,
            litchi_odraw::image::Limits::default(),
        )
        .map_err(PackageError::from)?;
        if resources
            .insert(
                id,
                PictureResource {
                    kind: image.kind(),
                    data,
                    blip,
                },
            )
            .is_some()
        {
            return Err(PackageError::Corrupted(
                "PPT image inventory contains duplicate BLIP IDs".into(),
            )
            .into());
        }
    }
    Ok(PictureInventory { slots, resources })
}

fn picture_remaps(
    target: &Snapshot,
    donor: &Snapshot,
    references: &BTreeSet<u32>,
) -> Result<(BTreeMap<u32, u32>, Vec<TransferredPicture>)> {
    if references.is_empty() {
        return Ok((BTreeMap::new(), Vec::new()));
    }
    let donor_pictures = picture_inventory(donor)?;
    let target_pictures = picture_inventory(target)?;
    let mut remaps = BTreeMap::new();
    let mut pictures = Vec::<TransferredPicture>::new();
    for source_id in references {
        let source = donor_pictures.resources.get(source_id).ok_or_else(|| {
            PackageError::Corrupted(format!(
                "PPT slide references missing donor BLIP {source_id}"
            ))
        })?;
        let existing = target_pictures
            .resources
            .iter()
            .find(|(_candidate_id, candidate)| {
                candidate.kind == source.kind && candidate.data == source.data
            })
            .map(|(candidate_id, _candidate)| *candidate_id);
        let staged = pictures.iter().find_map(|picture| {
            let blip = litchi_odraw::image::Blip::parse(&picture.blip, Default::default()).ok()?;
            (blip.kind() == source.kind && blip.data() == source.data).then_some(picture.target_id)
        });
        let target_id = if let Some(target_id) = existing.or(staged) {
            target_id
        } else {
            let target_id = target_pictures
                .slots
                .checked_add(u32::try_from(pictures.len()).map_err(|_error| {
                    PackageError::ResourceLimit("too many transferred PPT pictures".into())
                })?)
                .and_then(|value| value.checked_add(1))
                .ok_or_else(|| PackageError::ResourceLimit("PPT BLIP ID overflow".into()))?;
            litchi_odraw::image::Id::new(target_id).map_err(PackageError::from)?;
            build_picture_fbse(&source.blip, 0)?;
            pictures.push(TransferredPicture {
                target_id,
                blip: source.blip.clone(),
            });
            target_id
        };
        remaps.insert(*source_id, target_id);
    }
    Ok((remaps, pictures))
}

fn build_picture_fbse(blip_bytes: &[u8], delay_offset: u32) -> Result<Vec<u8>> {
    let blip = litchi_odraw::image::Blip::parse(blip_bytes, Default::default())
        .map_err(PackageError::from)?;
    let kind = blip.kind();
    let uid = blip
        .uids()
        .ok_or_else(|| {
            PackageError::InvalidFormat(
                "opaque OfficeArt BLIPs cannot be added to a PPT BStore".into(),
            )
        })?
        .effective();
    let size = u32::try_from(blip_bytes.len())
        .map_err(|_error| PackageError::ResourceLimit("PPT BLIP exceeds u32".into()))?;
    let mut fbse = Vec::with_capacity(44);
    fbse.extend_from_slice(&((u16::from(kind.raw()) << 4) | 2).to_le_bytes());
    fbse.extend_from_slice(&0xF007u16.to_le_bytes());
    fbse.extend_from_slice(&36u32.to_le_bytes());
    fbse.extend_from_slice(&[kind.raw(), kind.raw()]);
    fbse.extend_from_slice(uid.as_bytes());
    fbse.extend_from_slice(&0x00FFu16.to_le_bytes());
    fbse.extend_from_slice(&size.to_le_bytes());
    fbse.extend_from_slice(&1u32.to_le_bytes());
    fbse.extend_from_slice(&delay_offset.to_le_bytes());
    fbse.extend_from_slice(&[0, 0, 0, 0]);
    Ok(fbse)
}

fn append_transferred_pictures(
    snapshot: &Snapshot,
    document: &[u8],
    pictures: &BTreeMap<u32, Vec<u8>>,
) -> Result<(Vec<u8>, Vec<u8>)> {
    let inventory = picture_inventory(snapshot)?;
    let mut delay = picture_stream(snapshot)?;
    let mut entries = Vec::with_capacity(pictures.len());
    for (index, (target_id, blip)) in pictures.iter().enumerate() {
        let expected = inventory
            .slots
            .checked_add(u32::try_from(index).map_err(|_error| {
                PackageError::ResourceLimit("too many transferred PPT pictures".into())
            })?)
            .and_then(|value| value.checked_add(1))
            .ok_or_else(|| PackageError::ResourceLimit("PPT BLIP ID overflow".into()))?;
        if *target_id != expected {
            return Err(PackageError::Corrupted(
                "transferred PPT picture IDs are not contiguous after the target BStore".into(),
            )
            .into());
        }
        let offset = u32::try_from(delay.len()).map_err(|_error| {
            PackageError::ResourceLimit("PPT Pictures stream exceeds u32".into())
        })?;
        entries.push(build_picture_fbse(blip, offset)?);
        delay.extend_from_slice(blip);
    }

    let (mut root, consumed) = crate::Record::parse(document, 0, &snapshot.limits)?;
    if consumed != document.len() || root.record_type != crate::RecordType::Document {
        return Err(PackageError::Corrupted(
            "PPT picture-store owner is not one DocumentContainer".into(),
        )
        .into());
    }
    let mut owners = 0usize;
    append_picture_store_owner(&mut root, &entries, &mut owners)?;
    if owners != 1 {
        return Err(PackageError::Corrupted(
            "PPT document must contain exactly one PPDrawingGroup picture-store owner".into(),
        )
        .into());
    }
    Ok((encode_record(&root)?, delay))
}

fn append_picture_store_owner(
    record: &mut crate::Record,
    entries: &[Vec<u8>],
    owners: &mut usize,
) -> Result<()> {
    if record.record_type == crate::RecordType::PPDrawingGroup {
        record.data = append_officeart_bstore(&record.data, entries)?.into();
        *owners = owners.saturating_add(1);
    }
    for child in &mut record.children {
        append_picture_store_owner(child, entries, owners)?;
    }
    Ok(())
}

fn append_officeart_bstore(bytes: &[u8], entries: &[Vec<u8>]) -> Result<Vec<u8>> {
    let mut stores = 0usize;
    let rewritten = rewrite_officeart_bstore(bytes, entries, &mut stores)?;
    match stores {
        0 => {
            let mut containers = 0usize;
            let inserted = insert_officeart_bstore(&rewritten, entries, &mut containers)?;
            if containers != 1 {
                return Err(PackageError::Corrupted(
                    "PPDrawingGroup must contain exactly one OfficeArtDggContainer".into(),
                )
                .into());
            }
            Ok(inserted)
        },
        1 => Ok(rewritten),
        _ => Err(PackageError::Corrupted(
            "PPDrawingGroup contains multiple OfficeArt BStore containers".into(),
        )
        .into()),
    }
}

fn rewrite_officeart_bstore(
    bytes: &[u8],
    entries: &[Vec<u8>],
    stores: &mut usize,
) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    let mut offset = 0usize;
    while offset < bytes.len() {
        let (version, instance, kind, payload_start, end) =
            raw_record_bounds(bytes, offset, "OfficeArt drawing group")?;
        let mut payload = bytes[payload_start..end].to_vec();
        let mut rewritten_instance = instance;
        if kind == 0xF001 {
            if version != 0x0F {
                return Err(PackageError::Corrupted(
                    "OfficeArt BStore has an invalid record version".into(),
                )
                .into());
            }
            rewritten_instance = instance
                .checked_add(u16::try_from(entries.len()).map_err(|_error| {
                    PackageError::ResourceLimit("too many transferred PPT pictures".into())
                })?)
                .filter(|count| *count <= 0x0FFF)
                .ok_or_else(|| {
                    PackageError::ResourceLimit("PPT BStore exceeds 4095 entries".into())
                })?;
            for entry in entries {
                payload.extend_from_slice(entry);
            }
            *stores = stores.saturating_add(1);
        } else if version == 0x0F {
            payload = rewrite_officeart_bstore(&payload, entries, stores)?;
        }
        append_raw_record(&mut output, version, rewritten_instance, kind, &payload)?;
        offset = end;
    }
    Ok(output)
}

fn insert_officeart_bstore(
    bytes: &[u8],
    entries: &[Vec<u8>],
    containers: &mut usize,
) -> Result<Vec<u8>> {
    let bstore = encode_new_bstore(entries)?;
    let mut output = Vec::new();
    let mut offset = 0usize;
    while offset < bytes.len() {
        let (version, instance, kind, payload_start, end) =
            raw_record_bounds(bytes, offset, "OfficeArt drawing group")?;
        let source_payload = &bytes[payload_start..end];
        let payload = if kind == 0xF000 {
            *containers = containers.saturating_add(1);
            insert_bstore_after_dgg(source_payload, &bstore)?
        } else if version == 0x0F {
            insert_officeart_bstore(source_payload, entries, containers)?
        } else {
            source_payload.to_vec()
        };
        append_raw_record(&mut output, version, instance, kind, &payload)?;
        offset = end;
    }
    Ok(output)
}

fn insert_bstore_after_dgg(bytes: &[u8], bstore: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    let mut offset = 0usize;
    let mut inserted = false;
    while offset < bytes.len() {
        let (_version, _instance, kind, _payload_start, end) =
            raw_record_bounds(bytes, offset, "OfficeArtDggContainer")?;
        output.extend_from_slice(&bytes[offset..end]);
        if kind == 0xF006 && !inserted {
            output.extend_from_slice(bstore);
            inserted = true;
        }
        offset = end;
    }
    if !inserted {
        return Err(PackageError::Corrupted(
            "OfficeArtDggContainer has no OfficeArtDgg atom".into(),
        )
        .into());
    }
    Ok(output)
}

fn encode_new_bstore(entries: &[Vec<u8>]) -> Result<Vec<u8>> {
    let count = u16::try_from(entries.len())
        .ok()
        .filter(|count| *count <= 0x0FFF)
        .ok_or_else(|| PackageError::ResourceLimit("PPT BStore exceeds 4095 entries".into()))?;
    let mut payload = Vec::new();
    for entry in entries {
        payload.extend_from_slice(entry);
    }
    let mut output = Vec::new();
    append_raw_record(&mut output, 0x0F, count, 0xF001, &payload)?;
    Ok(output)
}

fn append_raw_record(
    output: &mut Vec<u8>,
    version: u16,
    instance: u16,
    kind: u16,
    payload: &[u8],
) -> Result<()> {
    let length = u32::try_from(payload.len())
        .map_err(|_error| PackageError::ResourceLimit("OfficeArt record exceeds u32".into()))?;
    output.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    output.extend_from_slice(&kind.to_le_bytes());
    output.extend_from_slice(&length.to_le_bytes());
    output.extend_from_slice(payload);
    Ok(())
}

fn picture_stream(snapshot: &Snapshot) -> Result<Vec<u8>> {
    let mut ole = litchi_cfb::OleFile::open(
        Cursor::new(snapshot.bytes()),
        litchi_cfb::OleFileLimits::default(),
    )
    .map_err(PackageError::from)?;
    let paths = ole
        .list_streams()
        .into_iter()
        .filter(|path| path.last().is_some_and(|name| name == "Pictures"))
        .collect::<Vec<_>>();
    match paths.as_slice() {
        [] => Ok(Vec::new()),
        [path] => {
            let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
            Ok(ole.open_stream(&refs).map_err(PackageError::from)?)
        },
        _ => Err(
            PackageError::Corrupted("PPT package contains multiple Pictures streams".into()).into(),
        ),
    }
}

fn validate_transferred_pictures(
    snapshot: &Snapshot,
    pictures: &BTreeMap<u32, Vec<u8>>,
) -> Result<()> {
    if pictures.is_empty() {
        return Ok(());
    }
    let inventory = picture_inventory(snapshot)?;
    for (target_id, expected_blip) in pictures {
        let expected = litchi_odraw::image::Blip::parse(expected_blip, Default::default())
            .map_err(PackageError::from)?;
        let actual = inventory.resources.get(target_id).ok_or_else(|| {
            PackageError::Corrupted(format!(
                "published PPT is missing transferred BLIP {target_id}"
            ))
        })?;
        if actual.kind != expected.kind() || actual.data != expected.data() {
            return Err(PackageError::Corrupted(format!(
                "published PPT transferred BLIP {target_id} changed semantic image bytes"
            ))
            .into());
        }
    }
    Ok(())
}

fn validate_unrelated_picture_publication(before: &[u8], after: &[u8]) -> Result<()> {
    let mut old =
        litchi_cfb::OleFile::open(Cursor::new(before), litchi_cfb::OleFileLimits::default())
            .map_err(PackageError::from)?;
    let mut new =
        litchi_cfb::OleFile::open(Cursor::new(after), litchi_cfb::OleFileLimits::default())
            .map_err(PackageError::from)?;
    let mut old_paths = old.list_streams();
    let mut new_paths = new.list_streams();
    old_paths.retain(|path| path.last().is_none_or(|name| name != "Pictures"));
    new_paths.retain(|path| path.last().is_none_or(|name| name != "Pictures"));
    old_paths.sort();
    new_paths.sort();
    if old_paths != new_paths {
        return Err(PackageError::Corrupted(
            "PPT picture publication changed unrelated CFB stream topology".into(),
        )
        .into());
    }
    for path in old_paths {
        if path.last().is_some_and(|name| {
            name.eq_ignore_ascii_case("PowerPoint Document")
                || name.eq_ignore_ascii_case("Current User")
        }) {
            continue;
        }
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        if old.open_stream(&refs).map_err(PackageError::from)?
            != new.open_stream(&refs).map_err(PackageError::from)?
        {
            return Err(PackageError::Corrupted(format!(
                "PPT picture publication changed unrelated stream {}",
                path.join("/")
            ))
            .into());
        }
    }
    Ok(())
}

fn drawing_remaps(target: &Snapshot, donor: &SlideRelationships) -> Result<BTreeMap<u32, u32>> {
    if donor
        .unsupported_dependencies
        .contains(&TransferDependency::AnimationBuildGraph)
    {
        return Err(Error::Refused(Refusal::UnsupportedSlideDependency {
            dependency: TransferDependency::AnimationBuildGraph,
        }));
    }
    if donor
        .unsupported_dependencies
        .contains(&TransferDependency::OtherOfficeArtShapeReference)
    {
        return Err(Error::Refused(Refusal::UnsupportedSlideDependency {
            dependency: TransferDependency::OtherOfficeArtShapeReference,
        }));
    }
    if !donor.connector_shape_ids.is_subset(&donor.shape_ids) {
        return Err(Error::Refused(Refusal::UnsupportedSlideDependency {
            dependency: TransferDependency::ConnectorRule,
        }));
    }
    if !donor.other_shape_reference_ids.is_subset(&donor.shape_ids) {
        return Err(Error::Refused(Refusal::UnsupportedSlideDependency {
            dependency: TransferDependency::OtherOfficeArtShapeReference,
        }));
    }
    if donor.shape_ids.is_empty() {
        return Ok(BTreeMap::new());
    }
    let live_persist_ids = target
        .document
        .slides()
        .iter()
        .map(|slide| slide.persist_id())
        .collect::<BTreeSet<_>>();
    let editor = crate::embedded::object::Editor::open_records_arc_with_limit(
        target.shared_bytes(),
        target.limits.max_package_bytes,
        target.limits,
    )?;
    let mut live_shape_ids = BTreeSet::new();
    let mut available_shape_ids = BTreeSet::new();
    for persist_id in editor.persist_ids() {
        let Ok(bytes) = editor.persisted_record(persist_id) else {
            continue;
        };
        let Ok((root, consumed)) = crate::Record::parse(&bytes, 0, &target.limits) else {
            continue;
        };
        if consumed != bytes.len() || root.record_type != crate::RecordType::Slide {
            continue;
        }
        let relationships = scan_slide_relationships(&root)?;
        if live_persist_ids.contains(&persist_id) {
            live_shape_ids.extend(relationships.shape_ids);
        } else {
            available_shape_ids.extend(relationships.shape_ids);
        }
    }
    available_shape_ids.retain(|shape_id| !live_shape_ids.contains(shape_id));
    let mut clusters = BTreeMap::<u32, BTreeSet<u32>>::new();
    for shape_id in available_shape_ids {
        clusters.entry(shape_id >> 10).or_default().insert(shape_id);
    }
    let mut remaining = clusters
        .into_values()
        .find(|cluster| cluster.len() >= donor.shape_ids.len())
        .ok_or(Error::Refused(Refusal::UnsupportedSlideDependency {
            dependency: TransferDependency::DrawingGroup,
        }))?;
    let mut remaps = BTreeMap::new();
    for source_id in &donor.shape_ids {
        let target_id = if remaining.remove(source_id) {
            *source_id
        } else {
            let candidate = remaining.iter().next().copied().ok_or(Error::Refused(
                Refusal::UnsupportedSlideDependency {
                    dependency: TransferDependency::DrawingGroup,
                },
            ))?;
            remaining.remove(&candidate);
            candidate
        };
        remaps.insert(*source_id, target_id);
    }
    Ok(remaps)
}

fn scan_officeart_records(bytes: &[u8], output: &mut SlideRelationships) -> Result<()> {
    visit_raw_records(
        bytes,
        "OfficeArt drawing",
        |version, instance, kind, payload| {
            match kind {
                0xF00A => {
                    output.shape_ids.insert(read_payload_u32(payload, 0)?);
                },
                0xF00B | 0xF121 | 0xF122 => {
                    output
                        .blip_ids
                        .extend(option_blip_references(instance, payload)?);
                    scan_linked_shape_references(instance, payload, output)?;
                },
                0xF012 if version == 1 && instance == 0 => {
                    scan_connector_rule(payload, output)?;
                },
                0xF012 => {
                    return Err(PackageError::Corrupted(
                        "OfficeArtFConnectorRule has an invalid record header".into(),
                    )
                    .into());
                },
                0xF014 if version == 0 && instance == 0 => {
                    scan_single_shape_rule(payload, "OfficeArtFArcRule", output)?;
                },
                0xF017 if version == 0 && instance == 0 => {
                    scan_single_shape_rule(payload, "OfficeArtFCalloutRule", output)?;
                },
                0xF11D if version == 0 && instance == 0 => {
                    scan_deleted_shape_rule(payload, output)?;
                },
                0xF119 if version == 0 && instance == 0 => {
                    scan_drawing_selection(payload, output)?;
                },
                0xF014 | 0xF017 | 0xF11D | 0xF119 => {
                    return Err(PackageError::Corrupted(
                        "OfficeArt shape-reference rule has an invalid record header".into(),
                    )
                    .into());
                },
                0xF00D | 0xF011 => scan_host_records(payload, output)?,
                _ if version == 0x0f => scan_officeart_records(payload, output)?,
                _ => {},
            }
            Ok(())
        },
    )
}

fn scan_host_records(bytes: &[u8], output: &mut SlideRelationships) -> Result<()> {
    visit_raw_records(
        bytes,
        "PowerPoint host records",
        |_version, _instance, kind, payload| {
            let record_type = crate::RecordType::from(kind);
            if record_type == crate::RecordType::InteractiveInfoAtom {
                scan_interactive_atom(payload, output)?;
            } else if record_type == crate::RecordType::ExternalObjectRefAtom {
                scan_external_reference(payload, output)?;
            } else if record_type == crate::RecordType::InteractiveInfo {
                scan_host_records(payload, output)?;
            } else if matches!(
                record_type,
                crate::RecordType::AnimationInfo
                    | crate::RecordType::BuildList
                    | crate::RecordType::ChartBuild
                    | crate::RecordType::DiagramBuild
                    | crate::RecordType::ParaBuild
            ) {
                output
                    .unsupported_dependencies
                    .insert(TransferDependency::AnimationBuildGraph);
            }
            Ok(())
        },
    )
}

fn scan_interactive_atom(bytes: &[u8], output: &mut SlideRelationships) -> Result<()> {
    let atom = crate::InteractiveInfoAtom::parse_payload(bytes)?;
    if atom.hyperlink_id != 0 {
        output.hyperlink_ids.insert(atom.hyperlink_id);
    }
    if atom.sound_id != 0 {
        output.sound_ids.insert(atom.sound_id);
    }
    if matches!(
        atom.action,
        crate::InteractionAction::Macro
            | crate::InteractionAction::RunProgram
            | crate::InteractionAction::Ole
    ) {
        output.active_action = true;
    }
    Ok(())
}

fn scan_external_reference(bytes: &[u8], output: &mut SlideRelationships) -> Result<()> {
    let id = read_payload_u32(bytes, 0)?;
    if bytes.len() != 4 || id == 0 {
        return Err(PackageError::Corrupted(
            "PPT ExternalObjectRefAtom has an invalid payload".into(),
        )
        .into());
    }
    output.external_object_ids.insert(id);
    Ok(())
}

fn hyperlink_remaps(
    target: &Snapshot,
    donor: &Snapshot,
    references: &BTreeSet<u32>,
) -> Result<BTreeMap<u32, u32>> {
    let donor_links = crate::Hyperlinks::parse(donor.document.record(), &donor.limits)?;
    let target_links = crate::Hyperlinks::parse(target.document.record(), &target.limits)?;
    references
        .iter()
        .map(|source_id| {
            let source = donor_links.get(*source_id).ok_or_else(|| {
                PackageError::Corrupted(format!(
                    "PPT slide references missing donor hyperlink {source_id}"
                ))
            })?;
            let target_id = target_links
                .hyperlinks
                .iter()
                .find(|candidate| hyperlink_semantics_equal(source, candidate))
                .map(|candidate| candidate.id)
                .ok_or(Error::Refused(Refusal::UnsupportedSlideDependency {
                    dependency: TransferDependency::HyperlinkAction,
                }))?;
            Ok((*source_id, target_id))
        })
        .collect()
}

fn hyperlink_semantics_equal(left: &crate::Hyperlink, right: &crate::Hyperlink) -> bool {
    left.friendly_name == right.friendly_name
        && left.target == right.target
        && left.location == right.location
        && left.extension == right.extension
}

fn sound_remaps(
    target: &Snapshot,
    donor: &Snapshot,
    references: &BTreeSet<u32>,
) -> Result<BTreeMap<u32, u32>> {
    if references.is_empty() {
        return Ok(BTreeMap::new());
    }
    let donor_owner =
        unique_record_owner(donor.document.record(), crate::RecordType::SoundCollection)?;
    let target_owner =
        unique_record_owner(target.document.record(), crate::RecordType::SoundCollection)?;
    let donor_sounds = crate::sound_collection::Collection::parse(donor_owner)?;
    let target_sounds = crate::sound_collection::Collection::parse(target_owner)?;
    references
        .iter()
        .map(|source_id| {
            let source = donor_sounds.get(*source_id).ok_or_else(|| {
                PackageError::Corrupted(format!(
                    "PPT slide references missing donor sound {source_id}"
                ))
            })?;
            let target_id = target_sounds
                .sounds
                .iter()
                .find(|candidate| sound_semantics_equal(source, candidate))
                .map(|candidate| candidate.id)
                .ok_or(Error::Refused(Refusal::UnsupportedSlideDependency {
                    dependency: TransferDependency::ExternalObjectRelationship,
                }))?;
            Ok((*source_id, target_id))
        })
        .collect()
}

fn sound_semantics_equal(
    left: &crate::sound_collection::Sound<'_>,
    right: &crate::sound_collection::Sound<'_>,
) -> bool {
    left.name == right.name
        && left.extension == right.extension
        && left.builtin_id == right.builtin_id
        && left.data == right.data
}

fn unique_record_owner(root: &crate::Record, owner: crate::RecordType) -> Result<&crate::Record> {
    let mut records = Vec::new();
    collect_record_owners(root, owner, &mut records);
    match records.as_slice() {
        [record] => Ok(*record),
        [] => Err(Error::Refused(Refusal::UnsupportedSlideDependency {
            dependency: TransferDependency::ExternalObjectRelationship,
        })),
        _ => Err(PackageError::Corrupted(format!(
            "PPT document contains multiple {owner:?} owners"
        ))
        .into()),
    }
}

fn external_object_remaps(
    target: &Snapshot,
    donor: &Snapshot,
    references: &BTreeSet<u32>,
) -> Result<BTreeMap<u32, u32>> {
    if references.is_empty() {
        return Ok(BTreeMap::new());
    }
    let donor_media = external_media_snapshot(donor)?;
    let target_media = external_media_snapshot(target)?;
    let donor_collection =
        donor_media
            .collection()
            .ok_or(Error::Refused(Refusal::UnsupportedSlideDependency {
                dependency: TransferDependency::ActiveOrUnknownExternalObject,
            }))?;
    let target_collection =
        target_media
            .collection()
            .ok_or(Error::Refused(Refusal::UnsupportedSlideDependency {
                dependency: TransferDependency::ExternalObjectRelationship,
            }))?;
    references
        .iter()
        .map(|source_id| {
            let source = donor_collection.get(*source_id).ok_or(Error::Refused(
                Refusal::UnsupportedSlideDependency {
                    dependency: TransferDependency::ActiveOrUnknownExternalObject,
                },
            ))?;
            let target_id = target_collection
                .objects
                .iter()
                .find(|candidate| media_semantics_equal(source, candidate))
                .map(crate::external_media::Object::id)
                .ok_or(Error::Refused(Refusal::UnsupportedSlideDependency {
                    dependency: TransferDependency::ExternalObjectRelationship,
                }))?;
            Ok((*source_id, target_id))
        })
        .collect()
}

fn media_semantics_equal(
    left: &crate::external_media::Object,
    right: &crate::external_media::Object,
) -> bool {
    use crate::external_media::Object;
    match (left, right) {
        (Object::Movie(left_movie), Object::Movie(right_movie)) => {
            left_movie.kind == right_movie.kind
                && left_movie.video.path == right_movie.video.path
                && media_values_equal(left_movie.video.media, right_movie.video.media)
        },
        (Object::LinkedAudio(left_audio), Object::LinkedAudio(right_audio)) => {
            left_audio.kind == right_audio.kind
                && left_audio.path == right_audio.path
                && media_values_equal(left_audio.media, right_audio.media)
        },
        (Object::CdAudio(left_cd), Object::CdAudio(right_cd)) => {
            left_cd.start == right_cd.start
                && left_cd.end == right_cd.end
                && media_values_equal(left_cd.media, right_cd.media)
        },
        (Object::EmbeddedWav(left_wav), Object::EmbeddedWav(right_wav)) => {
            left_wav.sound_id == right_wav.sound_id
                && left_wav.duration_ms == right_wav.duration_ms
                && media_values_equal(left_wav.media, right_wav.media)
        },
        _ => false,
    }
}

fn media_values_equal(
    left: crate::external_media::Media,
    right: crate::external_media::Media,
) -> bool {
    left.loop_playback == right.loop_playback
        && left.rewind_after_playing == right.rewind_after_playing
        && left.narration == right.narration
        && left.unused == right.unused
}

fn rewrite_slide_relationships(
    root: &mut crate::Record,
    hyperlinks: &BTreeMap<u32, u32>,
    sounds: &BTreeMap<u32, u32>,
    external_objects: &BTreeMap<u32, u32>,
    pictures: &BTreeMap<u32, u32>,
    drawings: &BTreeMap<u32, u32>,
) -> Result<()> {
    if root.record_type == crate::RecordType::InteractiveInfoAtom {
        rewrite_interactive_atom(&mut root.data, hyperlinks, sounds)?;
    } else if root.record_type == crate::RecordType::ExternalObjectRefAtom {
        rewrite_external_reference(&mut root.data, external_objects)?;
    } else if root.record_type == crate::RecordType::RoundTripShapeId12Atom {
        rewrite_u32_reference(&mut root.data, drawings, "RoundTripShapeId12Atom")?;
    } else if root.record_type == crate::RecordType::BinaryTagData {
        rewrite_host_records(
            &mut root.data,
            hyperlinks,
            sounds,
            external_objects,
            drawings,
        )?;
    } else if root.record_type == crate::RecordType::PPDrawing {
        rewrite_officeart_records(
            {
                rewrite_officeart_drawing_identity(&mut root.data, drawings)?;
                &mut root.data
            },
            hyperlinks,
            sounds,
            external_objects,
            pictures,
            drawings,
        )?;
    }
    for child in &mut root.children {
        rewrite_slide_relationships(
            child,
            hyperlinks,
            sounds,
            external_objects,
            pictures,
            drawings,
        )?;
    }
    Ok(())
}

fn rewrite_officeart_drawing_identity(
    bytes: &mut [u8],
    drawings: &BTreeMap<u32, u32>,
) -> Result<()> {
    if drawings.is_empty() {
        return Ok(());
    }
    let mut drawing_groups = drawings.values().map(|shape_id| shape_id >> 10);
    let drawing_group = drawing_groups
        .next()
        .filter(|drawing_group| *drawing_group != 0 && *drawing_group <= 0x0fff)
        .ok_or_else(|| PackageError::Corrupted("PPT target drawing ID is invalid".into()))?;
    if drawing_groups.any(|candidate| candidate != drawing_group) {
        return Err(PackageError::Corrupted(
            "PPT drawing remap spans multiple target drawing groups".into(),
        )
        .into());
    }
    let spid_cur = drawings
        .values()
        .copied()
        .max()
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| PackageError::Corrupted("PPT target shape ID overflow".into()))?;
    let mut rewritten = 0usize;
    rewrite_officeart_dg_records(
        bytes,
        u16::try_from(drawing_group)
            .map_err(|_error| PackageError::Corrupted("PPT drawing ID exceeds u16".into()))?,
        spid_cur,
        &mut rewritten,
    )?;
    if rewritten != 1 {
        return Err(PackageError::Corrupted(
            "PPT ordinary drawing must contain exactly one OfficeArtDg atom".into(),
        )
        .into());
    }
    Ok(())
}

fn rewrite_officeart_dg_records(
    bytes: &mut [u8],
    drawing_group: u16,
    spid_cur: u32,
    rewritten: &mut usize,
) -> Result<()> {
    let mut offset = 0usize;
    while offset < bytes.len() {
        let (version, _instance, kind, payload_start, end) =
            raw_record_bounds(bytes, offset, "OfficeArt drawing")?;
        if kind == 0xF008 {
            let payload = bytes.get_mut(payload_start..end).ok_or_else(|| {
                PackageError::Corrupted("OfficeArtDg payload is truncated".into())
            })?;
            if payload.len() != 8 {
                return Err(PackageError::Corrupted(
                    "OfficeArtDg payload must contain eight bytes".into(),
                )
                .into());
            }
            let header = (drawing_group << 4) | version;
            bytes[offset..offset + 2].copy_from_slice(&header.to_le_bytes());
            bytes[payload_start + 4..payload_start + 8].copy_from_slice(&spid_cur.to_le_bytes());
            *rewritten = rewritten.saturating_add(1);
        } else if version == 0x0f {
            rewrite_officeart_dg_records(
                &mut bytes[payload_start..end],
                drawing_group,
                spid_cur,
                rewritten,
            )?;
        }
        offset = end;
    }
    Ok(())
}

fn rewrite_officeart_records(
    bytes: &mut [u8],
    hyperlinks: &BTreeMap<u32, u32>,
    sounds: &BTreeMap<u32, u32>,
    external_objects: &BTreeMap<u32, u32>,
    pictures: &BTreeMap<u32, u32>,
    drawings: &BTreeMap<u32, u32>,
) -> Result<()> {
    visit_raw_records_mut(
        bytes,
        "OfficeArt drawing",
        |version, instance, kind, payload| {
            match kind {
                0xF00A => rewrite_u32_reference(payload, drawings, "OfficeArtSp.spid")?,
                0xF00B | 0xF121 | 0xF122 => {
                    rewrite_blip_references(instance, payload, pictures)?;
                    rewrite_linked_shape_references(instance, payload, drawings)?;
                },
                0xF012 if version == 1 && instance == 0 => {
                    rewrite_connector_rule(payload, drawings)?;
                },
                0xF012 => {
                    return Err(PackageError::Corrupted(
                        "OfficeArtFConnectorRule has an invalid record header".into(),
                    )
                    .into());
                },
                0xF014 if version == 0 && instance == 0 => {
                    rewrite_single_shape_rule(payload, drawings, "OfficeArtFArcRule")?;
                },
                0xF017 if version == 0 && instance == 0 => {
                    rewrite_single_shape_rule(payload, drawings, "OfficeArtFCalloutRule")?;
                },
                0xF11D if version == 0 && instance == 0 => {
                    rewrite_deleted_shape_rule(payload, drawings)?;
                },
                0xF119 if version == 0 && instance == 0 => {
                    rewrite_drawing_selection(payload, drawings)?;
                },
                0xF014 | 0xF017 | 0xF11D | 0xF119 => {
                    return Err(PackageError::Corrupted(
                        "OfficeArt shape-reference rule has an invalid record header".into(),
                    )
                    .into());
                },
                0xF00D | 0xF011 => {
                    rewrite_host_records(payload, hyperlinks, sounds, external_objects, drawings)?;
                },
                _ if version == 0x0f => {
                    rewrite_officeart_records(
                        payload,
                        hyperlinks,
                        sounds,
                        external_objects,
                        pictures,
                        drawings,
                    )?;
                },
                _ => {},
            }
            Ok(())
        },
    )
}

fn rewrite_linked_shape_references(
    instance: u16,
    payload: &mut [u8],
    drawings: &BTreeMap<u32, u32>,
) -> Result<()> {
    let property_bytes = usize::from(instance)
        .checked_mul(6)
        .ok_or_else(|| PackageError::Corrupted("OfficeArt property count overflow".into()))?;
    let properties = payload
        .get_mut(..property_bytes)
        .ok_or_else(|| PackageError::Corrupted("OfficeArt property table is truncated".into()))?;
    for property in properties.as_chunks_mut::<6>().0.iter_mut() {
        let opid = u16::from_le_bytes([property[0], property[1]]) & 0x3fff;
        if opid != 0x008a {
            continue;
        }
        let source_id = u32::from_le_bytes([property[2], property[3], property[4], property[5]]);
        if source_id == 0 {
            continue;
        }
        let target_id = drawings.get(&source_id).ok_or_else(|| {
            PackageError::Corrupted(format!(
                "OfficeArt hspNext references unmapped shape {source_id}"
            ))
        })?;
        property[2..6].copy_from_slice(&target_id.to_le_bytes());
    }
    Ok(())
}

fn rewrite_blip_references(
    instance: u16,
    payload: &mut [u8],
    pictures: &BTreeMap<u32, u32>,
) -> Result<()> {
    let property_bytes = usize::from(instance)
        .checked_mul(6)
        .ok_or_else(|| PackageError::Corrupted("OfficeArt property count overflow".into()))?;
    let properties = payload
        .get_mut(..property_bytes)
        .ok_or_else(|| PackageError::Corrupted("OfficeArt property table is truncated".into()))?;
    for property in properties.as_chunks_mut::<6>().0.iter_mut() {
        if u16::from_le_bytes([property[0], property[1]]) & 0x4000 == 0 {
            continue;
        }
        let source_id = u32::from_le_bytes([property[2], property[3], property[4], property[5]]);
        if source_id == 0 {
            continue;
        }
        let target_id = pictures.get(&source_id).ok_or_else(|| {
            PackageError::Corrupted(format!(
                "OfficeArt property references unmapped BLIP {source_id}"
            ))
        })?;
        property[2..6].copy_from_slice(&target_id.to_le_bytes());
    }
    Ok(())
}

fn rewrite_single_shape_rule(
    payload: &mut [u8],
    drawings: &BTreeMap<u32, u32>,
    context: &str,
) -> Result<()> {
    if payload.len() != 8 {
        return Err(
            PackageError::Corrupted(format!("{context} payload must contain eight bytes")).into(),
        );
    }
    rewrite_u32_reference(&mut payload[4..8], drawings, context)
}

fn rewrite_deleted_shape_rule(payload: &mut [u8], drawings: &BTreeMap<u32, u32>) -> Result<()> {
    if payload.len() != 4 {
        return Err(PackageError::Corrupted(
            "OfficeArtFPSPL payload must contain four bytes".into(),
        )
        .into());
    }
    let raw = read_payload_u32(payload, 0)?;
    let source_id = raw & 0x3fff_ffff;
    if source_id == 0 {
        return Ok(());
    }
    let target_id = drawings.get(&source_id).ok_or_else(|| {
        PackageError::Corrupted(format!(
            "OfficeArtFPSPL references unmapped shape {source_id}"
        ))
    })?;
    payload.copy_from_slice(&((raw & 0xc000_0000) | target_id).to_le_bytes());
    Ok(())
}

fn rewrite_drawing_selection(payload: &mut [u8], drawings: &BTreeMap<u32, u32>) -> Result<()> {
    if payload.len() < 12 || !(payload.len() - 12).is_multiple_of(4) {
        return Err(PackageError::Corrupted(
            "OfficeArtFDGSL payload has an invalid shape-list width".into(),
        )
        .into());
    }
    for offset in (8..payload.len()).step_by(4) {
        let source_id = read_payload_u32(payload, offset)?;
        if source_id == 0 {
            continue;
        }
        let target_id = drawings.get(&source_id).ok_or_else(|| {
            PackageError::Corrupted(format!(
                "OfficeArtFDGSL references unmapped shape {source_id}"
            ))
        })?;
        payload[offset..offset + 4].copy_from_slice(&target_id.to_le_bytes());
    }
    Ok(())
}

fn rewrite_connector_rule(payload: &mut [u8], drawings: &BTreeMap<u32, u32>) -> Result<()> {
    if payload.len() != 24 {
        return Err(PackageError::Corrupted(
            "OfficeArtFConnectorRule payload must contain 24 bytes".into(),
        )
        .into());
    }
    for offset in [4, 8, 12] {
        let source_id = read_payload_u32(payload, offset)?;
        let target_id = drawings.get(&source_id).ok_or_else(|| {
            PackageError::Corrupted(format!(
                "OfficeArtFConnectorRule references unmapped shape {source_id}"
            ))
        })?;
        payload[offset..offset + 4].copy_from_slice(&target_id.to_le_bytes());
    }
    Ok(())
}

fn rewrite_host_records(
    bytes: &mut [u8],
    hyperlinks: &BTreeMap<u32, u32>,
    sounds: &BTreeMap<u32, u32>,
    external_objects: &BTreeMap<u32, u32>,
    drawings: &BTreeMap<u32, u32>,
) -> Result<()> {
    visit_raw_records_mut(
        bytes,
        "PowerPoint host records",
        |_version, _instance, kind, payload| {
            let record_type = crate::RecordType::from(kind);
            if record_type == crate::RecordType::InteractiveInfoAtom {
                rewrite_interactive_atom(payload, hyperlinks, sounds)?;
            } else if record_type == crate::RecordType::ExternalObjectRefAtom {
                rewrite_external_reference(payload, external_objects)?;
            } else if record_type == crate::RecordType::RoundTripShapeId12Atom {
                rewrite_u32_reference(payload, drawings, "RoundTripShapeId12Atom")?;
            } else if record_type == crate::RecordType::InteractiveInfo {
                rewrite_host_records(payload, hyperlinks, sounds, external_objects, drawings)?;
            }
            Ok(())
        },
    )
}

fn rewrite_interactive_atom(
    bytes: &mut [u8],
    hyperlinks: &BTreeMap<u32, u32>,
    sounds: &BTreeMap<u32, u32>,
) -> Result<()> {
    let mut atom = crate::InteractiveInfoAtom::parse_payload(bytes)?;
    if let Some(id) = sounds.get(&atom.sound_id) {
        atom.sound_id = *id;
    }
    if let Some(id) = hyperlinks.get(&atom.hyperlink_id) {
        atom.hyperlink_id = *id;
    }
    bytes.copy_from_slice(&atom.to_payload());
    Ok(())
}

fn rewrite_external_reference(bytes: &mut [u8], remaps: &BTreeMap<u32, u32>) -> Result<()> {
    let id = read_payload_u32(bytes, 0)?;
    if bytes.len() != 4 || id == 0 {
        return Err(PackageError::Corrupted(
            "PPT ExternalObjectRefAtom has an invalid payload".into(),
        )
        .into());
    }
    if let Some(replacement) = remaps.get(&id) {
        bytes.copy_from_slice(&replacement.to_le_bytes());
    }
    Ok(())
}

fn rewrite_u32_reference(
    bytes: &mut [u8],
    remaps: &BTreeMap<u32, u32>,
    context: &str,
) -> Result<()> {
    if bytes.len() < 4 {
        return Err(
            PackageError::Corrupted(format!("PPT {context} reference is truncated")).into(),
        );
    }
    let id = read_payload_u32(bytes, 0)?;
    if let Some(replacement) = remaps.get(&id) {
        bytes[0..4].copy_from_slice(&replacement.to_le_bytes());
    }
    Ok(())
}

fn visit_raw_records(
    bytes: &[u8],
    context: &str,
    mut visit: impl FnMut(u16, u16, u16, &[u8]) -> Result<()>,
) -> Result<()> {
    let mut offset = 0usize;
    while offset < bytes.len() {
        let (version, instance, kind, payload_start, end) =
            raw_record_bounds(bytes, offset, context)?;
        visit(version, instance, kind, &bytes[payload_start..end])?;
        offset = end;
    }
    Ok(())
}

fn visit_raw_records_mut(
    bytes: &mut [u8],
    context: &str,
    mut visit: impl FnMut(u16, u16, u16, &mut [u8]) -> Result<()>,
) -> Result<()> {
    let mut offset = 0usize;
    while offset < bytes.len() {
        let (version, instance, kind, payload_start, end) =
            raw_record_bounds(bytes, offset, context)?;
        visit(version, instance, kind, &mut bytes[payload_start..end])?;
        offset = end;
    }
    Ok(())
}

fn raw_record_bounds(
    bytes: &[u8],
    offset: usize,
    context: &str,
) -> Result<(u16, u16, u16, usize, usize)> {
    let header_end = offset
        .checked_add(8)
        .ok_or_else(|| PackageError::Corrupted(format!("{context} header offset overflow")))?;
    let header = bytes
        .get(offset..header_end)
        .ok_or_else(|| PackageError::Corrupted(format!("truncated record header in {context}")))?;
    let version_instance = u16::from_le_bytes([header[0], header[1]]);
    let kind = u16::from_le_bytes([header[2], header[3]]);
    let length = usize::try_from(u32::from_le_bytes([
        header[4], header[5], header[6], header[7],
    ]))
    .map_err(|_error| PackageError::Corrupted(format!("{context} record size overflow")))?;
    let end = header_end
        .checked_add(length)
        .filter(|end| *end <= bytes.len())
        .ok_or_else(|| PackageError::Corrupted(format!("record extends beyond {context}")))?;
    Ok((
        version_instance & 0x0f,
        version_instance >> 4,
        kind,
        header_end,
        end,
    ))
}

fn collect_record_owners<'a>(
    record: &'a crate::Record,
    owner: crate::RecordType,
    output: &mut Vec<&'a crate::Record>,
) {
    if record.record_type == owner {
        output.push(record);
    }
    for child in &record.children {
        collect_record_owners(child, owner, output);
    }
}

fn external_owner_contains_active_or_unknown(root: &crate::Record) -> bool {
    let mut owners = Vec::new();
    collect_record_owners(root, crate::RecordType::ExObjList, &mut owners);
    owners.into_iter().any(contains_active_or_unknown_external)
}

fn contains_active_or_unknown_external(record: &crate::Record) -> bool {
    matches!(
        record.record_type,
        crate::RecordType::Unknown
            | crate::RecordType::ExternalOleEmbed
            | crate::RecordType::ExternalOleEmbedAtom
            | crate::RecordType::ExternalOleLink
            | crate::RecordType::ExternalOleLinkAtom
            | crate::RecordType::ExternalOleControl
            | crate::RecordType::ExternalOleControlAtom
    ) || record
        .children
        .iter()
        .any(contains_active_or_unknown_external)
}

const fn transfer_dependency_order(dependency: TransferDependency) -> u8 {
    match dependency {
        TransferDependency::SpeakerNotes => 0,
        TransferDependency::CommentAuthorCatalog => 1,
        TransferDependency::HyperlinkAction => 2,
        TransferDependency::DrawingGroup => 3,
        TransferDependency::ConnectorRule => 4,
        TransferDependency::OtherOfficeArtShapeReference => 5,
        TransferDependency::PictureStore => 6,
        TransferDependency::AnimationBuildGraph => 7,
        TransferDependency::ExternalObjectRelationship => 8,
        TransferDependency::ActiveOrUnknownExternalObject => 9,
    }
}

fn normalized_payload(payload: &SlidePayload) -> SlidePayload {
    let mut normalized = payload.clone();
    if let Some(atom) = normalized.group.first_mut()
        && atom.data.len() == 20
    {
        atom.data[0..4].fill(0);
        atom.data[12..16].fill(0);
    }
    normalized
}

fn inverse_structural(structural_change: &StructuralChange) -> StructuralChange {
    match structural_change {
        StructuralChange::Move(move_change) => StructuralChange::Move(move_change.inverse()),
        StructuralChange::Remove(list_change) => StructuralChange::Insert(ListChange {
            position: list_change.position,
            before_order: list_change.after_order,
            after_order: list_change.before_order,
            payload: normalized_payload(&list_change.payload),
        }),
        StructuralChange::Insert(list_change) => StructuralChange::Remove(ListChange {
            position: list_change.position,
            before_order: list_change.after_order,
            after_order: list_change.before_order,
            payload: normalized_payload(&list_change.payload),
        }),
        StructuralChange::Document(change, order) => {
            StructuralChange::Document(change.inverse(), *order)
        },
    }
}

fn replay_structural_changes(
    document: &mut document_structure::Transaction,
    changes: &[StructuralChange],
) -> Result<()> {
    for structural_change in changes {
        match structural_change {
            StructuralChange::Move(move_change) => {
                document.move_slide(move_change.from.get(), move_change.destination.get())?;
            },
            StructuralChange::Remove(remove_change) => {
                let removed = document.remove_slide(remove_change.position.get())?;
                if removed != remove_change.payload.group {
                    return Err(PackageError::Corrupted(
                        "PPT structural rebase selected a different slide-list group".into(),
                    )
                    .into());
                }
            },
            StructuralChange::Insert(insert_change) => {
                document.insert_slide_group(
                    insert_change.position.get(),
                    insert_change.payload.group.clone(),
                )?;
            },
            StructuralChange::Document(change, _) => change.replay(document)?,
        }
    }
    Ok(())
}

fn patch_effects(patch: &Patch) -> BTreeSet<String> {
    let mut effects = BTreeSet::new();
    for change in &patch.visibility_changes {
        effects.insert(format!("slide-id:{}/hidden", change.slide_id));
    }
    for change in &patch.advance_changes {
        effects.insert(format!("slide-id:{}/advance", change.slide_id));
    }
    for change in &patch.transition_visual_changes {
        effects.insert(format!("slide-id:{}/transition-visual", change.slide_id));
    }
    for change in &patch.text_changes {
        effects.insert(format!(
            "slide:{}/shape:{}/text",
            change.target.slide().get(),
            change.target.shape().get()
        ));
    }
    for change in &patch.anchor_changes {
        effects.insert(format!(
            "slide:{}/shape:{}/anchor",
            change.target.slide().get(),
            change.target.shape().get()
        ));
    }
    for change in &patch.media_path_changes {
        effects.insert(format!("external-media:{}/path", change.id));
    }
    for change in &patch.media_playback_changes {
        effects.insert(format!("external-media:{}/playback", change.id));
    }
    for structural_change in &patch.structural {
        match structural_change {
            StructuralChange::Move(move_change) => {
                let start = move_change.from.get().min(move_change.destination.get());
                let end = move_change.from.get().max(move_change.destination.get());
                for position in start..=end {
                    effects.insert(format!("slide-list/position:{position}"));
                }
            },
            StructuralChange::Remove(list_change) | StructuralChange::Insert(list_change) => {
                effects.insert(format!(
                    "slide-list/position:{}",
                    list_change.position.get()
                ));
                if let Ok((root, consumed)) =
                    crate::Record::parse(&list_change.payload.record, 0, &SourceLimits::default())
                    && consumed == list_change.payload.record.len()
                    && let Ok(relationships) = scan_slide_relationships(&root)
                {
                    effects.extend(
                        relationships
                            .shape_ids
                            .into_iter()
                            .map(|shape_id| format!("drawing/shape-id:{shape_id}")),
                    );
                }
            },
            StructuralChange::Document(change, _) => {
                effects.insert(change.effect());
            },
        }
    }
    effects
}

fn encode_payload(payload: &SlidePayload) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(if payload.pictures.is_empty() {
        b"LSP1"
    } else {
        b"LSP2"
    });
    bytes.extend_from_slice(&payload.master_id.to_le_bytes());
    bytes.extend_from_slice(
        &u32::try_from(payload.group.len())
            .map_err(|_error| PackageError::ResourceLimit("too many slide group records".into()))?
            .to_le_bytes(),
    );
    for item in &payload.group {
        let encoded = encode_record(item)?;
        append_length_prefixed(&mut bytes, &encoded)?;
    }
    append_length_prefixed(&mut bytes, &payload.record)?;
    if !payload.pictures.is_empty() {
        bytes.extend_from_slice(
            &u32::try_from(payload.pictures.len())
                .map_err(|_error| {
                    PackageError::ResourceLimit("too many transferred PPT pictures".into())
                })?
                .to_le_bytes(),
        );
        for picture in &payload.pictures {
            bytes.extend_from_slice(&picture.target_id.to_le_bytes());
            append_length_prefixed(&mut bytes, &picture.blip)?;
        }
    }
    Ok(bytes)
}

fn decode_payload(bytes: &[u8], limits: SourceLimits) -> Result<SlidePayload> {
    let version = bytes.get(0..4);
    if bytes.len() > limits.max_input_bytes || !matches!(version, Some(b"LSP1" | b"LSP2")) {
        return Err(invalid_durable_patch(
            "invalid or oversized slide insertion blob",
        ));
    }
    let master_id = read_payload_u32(bytes, 4)?;
    let count = usize::try_from(read_payload_u32(bytes, 8)?)
        .map_err(|_error| invalid_durable_patch("slide group count exceeds this platform"))?;
    if count == 0 || count > limits.max_records {
        return Err(invalid_durable_patch("slide group count exceeds its bound"));
    }
    let mut offset = 12usize;
    let mut group = Vec::with_capacity(count);
    for _ in 0..count {
        let encoded = take_length_prefixed(bytes, &mut offset, limits.max_record_bytes)?;
        let (record, consumed) = crate::Record::parse(encoded, 0, &limits)?;
        if consumed != encoded.len() {
            return Err(invalid_durable_patch(
                "slide group blob record has trailing bytes",
            ));
        }
        group.push(record);
    }
    let record = take_length_prefixed(bytes, &mut offset, limits.max_record_bytes)?.to_vec();
    let mut pictures = Vec::new();
    if version == Some(b"LSP2") {
        let picture_count = usize::try_from(read_payload_u32(bytes, offset)?)
            .map_err(|_error| invalid_durable_patch("picture count exceeds this platform"))?;
        offset = offset
            .checked_add(4)
            .ok_or_else(|| invalid_durable_patch("picture count offset overflows"))?;
        if picture_count == 0 || picture_count > 0x0FFF {
            return Err(invalid_durable_patch(
                "transferred picture count exceeds its bound",
            ));
        }
        pictures.reserve(picture_count);
        for _ in 0..picture_count {
            let target_id = read_payload_u32(bytes, offset)?;
            offset = offset
                .checked_add(4)
                .ok_or_else(|| invalid_durable_patch("picture ID offset overflows"))?;
            litchi_odraw::image::Id::new(target_id)
                .map_err(|_error| invalid_durable_patch("transferred picture ID is invalid"))?;
            let blip = take_length_prefixed(bytes, &mut offset, limits.max_record_bytes)?.to_vec();
            build_picture_fbse(&blip, 0)
                .map_err(|_error| invalid_durable_patch("transferred picture BLIP is invalid"))?;
            pictures.push(TransferredPicture { target_id, blip });
        }
    }
    if offset != bytes.len() {
        return Err(invalid_durable_patch(
            "slide insertion blob has trailing bytes",
        ));
    }
    let mut payload = slide_payload(group, record)?;
    payload.pictures = pictures;
    if payload.master_id != master_id {
        return Err(invalid_durable_patch(
            "slide insertion blob master is inconsistent",
        ));
    }
    Ok(payload)
}

fn encode_record(record: &crate::Record) -> Result<Vec<u8>> {
    let mut payload = Vec::new();
    if record.children.is_empty() {
        payload.extend_from_slice(&record.data);
    } else {
        for child in &record.children {
            payload.extend_from_slice(&encode_record(child)?);
        }
    }
    let length = u32::try_from(payload.len()).map_err(|_error| {
        PackageError::ResourceLimit("slide transfer record exceeds u32".into())
    })?;
    let mut bytes = Vec::with_capacity(8usize.saturating_add(payload.len()));
    bytes.extend_from_slice(&(record.version | (record.instance << 4)).to_le_bytes());
    bytes.extend_from_slice(&record.record_type_raw.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(&payload);
    Ok(bytes)
}

fn append_length_prefixed(target: &mut Vec<u8>, payload: &[u8]) -> Result<()> {
    let length = u32::try_from(payload.len()).map_err(|_error| {
        PackageError::ResourceLimit("slide transfer payload exceeds u32".into())
    })?;
    target.extend_from_slice(&length.to_le_bytes());
    target.extend_from_slice(payload);
    Ok(())
}

fn take_length_prefixed<'a>(bytes: &'a [u8], offset: &mut usize, limit: usize) -> Result<&'a [u8]> {
    let length = usize::try_from(read_payload_u32(bytes, *offset)?)
        .map_err(|_error| invalid_durable_patch("slide blob length exceeds this platform"))?;
    *offset = offset
        .checked_add(4)
        .ok_or_else(|| invalid_durable_patch("slide blob offset overflow"))?;
    if length > limit {
        return Err(invalid_durable_patch("slide blob item exceeds its bound"));
    }
    let end = offset
        .checked_add(length)
        .ok_or_else(|| invalid_durable_patch("slide blob range overflow"))?;
    let payload = bytes
        .get(*offset..end)
        .ok_or_else(|| invalid_durable_patch("slide blob item is truncated"))?;
    *offset = end;
    Ok(payload)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use litchi_cfb::OleError;
    use litchi_cfb::{OleFile, OleWriter};
    use litchi_core::{ReadAt, SourceVersion};
    use std::sync::atomic::{AtomicU64, Ordering};

    fn fixture(name: &str) -> Vec<u8> {
        std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/poi/test-data/slideshow")
                .join(name),
        )
        .unwrap()
    }

    fn slide_texts(bytes: &[u8]) -> Vec<String> {
        let mut package = Reader::from_reader(Cursor::new(bytes)).unwrap();
        package
            .presentation()
            .unwrap()
            .slides()
            .unwrap()
            .iter()
            .map(|slide| slide.text().unwrap().to_string())
            .collect()
    }

    fn authored_fixture() -> Vec<u8> {
        let mut writer = crate::writer::Builder::new();
        let first = writer.add_slide().unwrap();
        writer
            .add_textbox(first, 10, 10, 240, 40, "first slide")
            .unwrap();
        let second = writer.add_slide().unwrap();
        writer
            .add_textbox(second, 10, 10, 240, 40, "second slide")
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        crate::raw::encode(&mut writer, &mut output).unwrap();
        output.into_inner()
    }

    #[test]
    fn package_backed_snapshot_matches_a_fresh_parse() {
        let bytes = authored_fixture();
        let limits = SourceLimits::builder()
            .max_input_bytes(bytes.len())
            .unwrap()
            .max_events(10_000)
            .unwrap()
            .build()
            .unwrap();
        let options = OpenOptions::new()
            .source_limits(limits)
            .leniency(crate::leniency::Leniency::TolerateUserEditChainDefects);
        let package = crate::Package::from_bytes_with_options(&bytes, options).unwrap();
        let artifact = Arc::from(package.to_encrypted_bytes().unwrap().into_boxed_slice());
        let admitted =
            Snapshot::from_package_with_options(&package, Arc::clone(&artifact), options).unwrap();
        let reparsed = Snapshot::from_bytes_with_options(bytes, options).unwrap();

        assert_eq!(admitted.bytes(), reparsed.bytes());
        assert_eq!(admitted.document, reparsed.document);
        assert_eq!(admitted.document_persist_id, reparsed.document_persist_id);
        assert_eq!(admitted.chain_repaired, reparsed.chain_repaired);
        assert_eq!(admitted.has_review_history, reparsed.has_review_history);
        assert_eq!(admitted.has_signature, reparsed.has_signature);
        assert!(Arc::ptr_eq(admitted.artifact.shared(), &artifact));
        assert_eq!(admitted.limits(), limits);
        assert_eq!(
            admitted.leniency(),
            crate::leniency::Leniency::TolerateUserEditChainDefects
        );
    }

    struct VersionedSource {
        bytes: Vec<u8>,
        revision: AtomicU64,
    }

    impl VersionedSource {
        fn new(bytes: Vec<u8>) -> Self {
            Self {
                bytes,
                revision: AtomicU64::new(0),
            }
        }

        fn change(&self) {
            self.revision.fetch_add(1, Ordering::AcqRel);
        }
    }

    impl ReadAt for VersionedSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(
                0x5050,
                self.revision.load(Ordering::Acquire),
            ))
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            let offset = usize::try_from(offset)
                .map_err(|_| std::io::Error::from(std::io::ErrorKind::InvalidInput))?;
            let input = self.bytes.get(offset..).unwrap_or_default();
            let count = input.len().min(output.len());
            output[..count].copy_from_slice(&input[..count]);
            Ok(count)
        }
    }

    #[test]
    fn package_backed_snapshot_rechecks_a_changed_positional_source() {
        let bytes = authored_fixture();
        let limits = SourceLimits::builder()
            .max_input_bytes(bytes.len())
            .unwrap()
            .max_events(10_000)
            .unwrap()
            .build()
            .unwrap();
        let options = OpenOptions::new()
            .source_limits(limits)
            .leniency(crate::leniency::Leniency::TolerateUserEditChainDefects);
        let source = Arc::new(VersionedSource::new(bytes));
        let package = crate::Package::from_read_at_with_options(source.clone(), options).unwrap();
        let artifact = Arc::from(package.to_encrypted_bytes().unwrap().into_boxed_slice());

        // Populate the package's signed-component memo before changing the
        // source. The freshness fence must still reject a later projection;
        // a cached signature result cannot make a positional source current.
        Snapshot::from_package_with_options(&package, Arc::clone(&artifact), options).unwrap();
        source.change();

        assert!(matches!(
            Snapshot::from_package_with_options(&package, artifact, options),
            Err(Error::Package(PackageError::Ole(
                OleError::SourceChanged { .. }
            )))
        ));
    }

    fn signed_fixture(bytes: &[u8]) -> Vec<u8> {
        let mut source =
            OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
        let mut writer = OleWriter::new();
        for path in source.list_streams() {
            let names: Vec<&str> = path.iter().map(String::as_str).collect();
            let data = source.open_stream(&names).unwrap();
            writer.create_stream(&names, &data).unwrap();
        }
        writer
            .create_stream(&["_signatures"], b"opaque synthetic signature")
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn authored_batch_fixture() -> Vec<u8> {
        let mut writer = crate::writer::Builder::new();
        let first = writer.add_slide().unwrap();
        writer
            .add_textbox(first, 10, 10, 240, 40, "first shape")
            .unwrap();
        writer
            .add_textbox(first, 10, 60, 240, 40, "second shape")
            .unwrap();
        let second = writer.add_slide().unwrap();
        writer
            .add_textbox(second, 10, 10, 240, 40, "third shape")
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        crate::raw::encode(&mut writer, &mut output).unwrap();
        output.into_inner()
    }

    fn authored_many_shapes_fixture(count: usize) -> Vec<u8> {
        let mut writer = crate::writer::Builder::new();
        let slide = writer.add_slide().unwrap();
        for index in 0..count {
            let text = format!("shape {index}");
            writer.add_textbox(slide, 10, 10, 240, 40, &text).unwrap();
        }
        let mut output = Cursor::new(Vec::new());
        crate::raw::encode(&mut writer, &mut output).unwrap();
        output.into_inner()
    }

    fn shape_text(bytes: &[u8], target: crate::text_edit::Target) -> String {
        let mut package = Reader::from_reader(Cursor::new(bytes)).unwrap();
        let presentation = package.presentation().unwrap();
        presentation.slides().unwrap()[target.slide().get()]
            .shapes()
            .unwrap()[target.shape().get()]
        .text()
        .unwrap()
        .to_string()
    }

    #[test]
    fn open_ole_live_document_inspection_matches_byte_ingress() {
        let bytes = authored_fixture();
        let expected = crate::embedded::object::Editor::inspect_live_document(&bytes).unwrap();
        let mut ole = OleFile::open(
            Cursor::new(bytes.as_slice()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        let actual =
            crate::embedded::object::Editor::inspect_live_document_from_ole(&mut ole).unwrap();

        assert_eq!(actual, expected);
        assert_eq!(
            Snapshot::from_bytes(bytes).unwrap().document_persist_id,
            actual.0
        );
    }

    fn authored_table_fixture() -> Vec<u8> {
        let mut writer = crate::writer::Builder::new();
        let slide = writer.add_slide().unwrap();
        let mut table = crate::writer::Table::new(2, 2).unwrap();
        table.set_cell_text(0, 0, "A1").unwrap();
        table.set_cell_text(1, 1, "B2").unwrap();
        writer.add_table(slide, 50, 60, table).unwrap();
        let mut output = Cursor::new(Vec::new());
        crate::raw::encode(&mut writer, &mut output).unwrap();
        output.into_inner()
    }

    fn anchor_target(snapshot: &Snapshot) -> crate::text_edit::Target {
        (0..snapshot.slide_count())
            .flat_map(|slide| {
                (0..64).map(move |shape| {
                    crate::text_edit::Target::new(Position::new(slide), Position::new(shape))
                })
            })
            .find(|target| {
                crate::text_edit::inspect_shape_anchor(snapshot.bytes(), *target).is_ok()
            })
            .expect("fixture must expose one anchored shape")
    }

    fn real_media_fixture() -> (Snapshot, u32) {
        for name in [
            "sound.ppt",
            "WithLinks.ppt",
            "ppt_with_embeded.ppt",
            "ole2-embedding-2003.ppt",
        ] {
            let Ok(snapshot) = Snapshot::from_bytes(fixture(name)) else {
                continue;
            };
            let Ok(media) = external_media_snapshot(&snapshot) else {
                continue;
            };
            if let Some(id) = media
                .collection()
                .and_then(|collection| collection.objects.first())
                .map(crate::external_media::Object::id)
            {
                return (snapshot, id);
            }
        }
        panic!("real fixture must expose one external-media object");
    }

    fn genuine_advance_cases() -> Vec<(&'static str, &'static str, Snapshot, Position)> {
        let roots = [
            (
                "poi-slideshow",
                std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                    .join("../../test-data/poi/test-data/slideshow"),
                &[
                    "basic_test_ppt_file.ppt",
                    "SampleShow.ppt",
                    "WithLinks.ppt",
                    "sound.ppt",
                    "45543.ppt",
                    "incorrect_slide_order.ppt",
                ][..],
            ),
            (
                "ole-ppt",
                std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/ppt"),
                &[
                    "SampleShow.ppt",
                    "text_shapes.ppt",
                    "text-margins.ppt",
                    "ppt_with_png.ppt",
                    "empty_textbox.ppt",
                ][..],
            ),
        ];
        let mut cases = Vec::new();
        for (family, root, names) in roots {
            for name in names {
                let Ok(bytes) = std::fs::read(root.join(name)) else {
                    continue;
                };
                let Ok(snapshot) = Snapshot::from_bytes(bytes) else {
                    continue;
                };
                for index in 0..snapshot.slide_count() {
                    let position = Position::new(index);
                    if snapshot.slide_advance(position).is_ok() {
                        cases.push((family, *name, snapshot, position));
                        break;
                    }
                }
            }
        }
        cases
    }

    fn advance_atom_data(snapshot: &Snapshot, position: Position) -> [u8; 16] {
        let slide = snapshot.document.slides()[position.get()];
        let bytes = persisted_record(snapshot, slide.persist_id()).unwrap();
        let (root, consumed) = crate::Record::parse(&bytes, 0, &snapshot.limits).unwrap();
        assert_eq!(consumed, bytes.len());
        unique_slide_advance_atom(&root, position)
            .unwrap()
            .data
            .as_slice()
            .try_into()
            .unwrap()
    }

    fn patch_limits() -> litchi_core::patch::PatchLimits {
        litchi_core::patch::PatchLimits::new(
            litchi_core::patch::BlobLimits::new(0, 0, 0),
            64 * 1024,
            64,
            8,
            4096,
            32 * 1024,
        )
    }

    fn transfer_patch_limits() -> litchi_core::patch::PatchLimits {
        litchi_core::patch::PatchLimits::new(
            litchi_core::patch::BlobLimits::new(8, 8 * 1024 * 1024, 16 * 1024 * 1024),
            20 * 1024 * 1024,
            64,
            8,
            4096,
            18 * 1024 * 1024,
        )
    }

    #[test]
    fn real_fixture_reorders_reopens_and_round_trips_durable_undo_redo() {
        let bytes = fixture("basic_test_ppt_file.ppt");
        let original_text = slide_texts(&bytes);
        let source = Snapshot::from_bytes(bytes).unwrap();
        let mut edit = source.edit().unwrap();
        edit.move_slide(Position::new(0), Position::new(1)).unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(
            slide_texts(commit.snapshot().bytes()),
            [original_text[1].clone(), original_text[0].clone(),]
        );
        assert_eq!(slide_texts(source.bytes()), original_text);
        assert_eq!(commit.patch().changes().len(), 1);
        assert_eq!(commit.patch().apply(&source).unwrap(), *commit.snapshot());
        assert_eq!(
            commit.patch().inverse().apply(commit.snapshot()).unwrap(),
            source
        );

        let durable = commit.patch().to_durable(patch_limits()).unwrap();
        let wire = durable.to_deterministic_json().unwrap();
        let decoded =
            litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
                &wire,
                patch_limits(),
            )
            .unwrap();
        assert_eq!(source.apply_durable(&decoded).unwrap(), *commit.snapshot());
        let durably_restored = commit.snapshot().apply_durable(&decoded.inverse()).unwrap();
        assert_eq!(slide_texts(durably_restored.bytes()), original_text);

        // Undo and redo are the inverse and the patch applied again.
        let undone = commit.patch().inverse().apply(commit.snapshot()).unwrap();
        assert_eq!(undone, source);
        assert_eq!(commit.patch().apply(&undone).unwrap(), *commit.snapshot());
    }

    #[test]
    fn signed_durable_semantic_noop_replays_exactly_and_changed_replay_is_refused() {
        use litchi_core::patch::{BlobBundle, Reversible, ReversibleOperation};

        let source = Snapshot::from_bytes(signed_fixture(&authored_fixture())).unwrap();
        let target = crate::text_edit::Target::new(Position::new(0), Position::new(0));
        let limits = patch_limits();
        let no_op = litchi_core::patch::Patch::<Reversible>::new(
            limits,
            "litchi-ppt",
            [ReversibleOperation::new(
                durable_text_operation(limits, target, "first slide", "first slide").unwrap(),
                durable_text_operation(limits, target, "first slide", "first slide").unwrap(),
            )],
            BlobBundle::new(limits.blobs()),
            BlobBundle::new(limits.blobs()),
        )
        .unwrap();
        let replayed = source.apply_durable(&no_op).unwrap();
        assert_eq!(replayed.bytes(), source.bytes());

        let changed = litchi_core::patch::Patch::<Reversible>::new(
            limits,
            "litchi-ppt",
            [ReversibleOperation::new(
                durable_text_operation(limits, target, "first slide", "changed").unwrap(),
                durable_text_operation(limits, target, "changed", "first slide").unwrap(),
            )],
            BlobBundle::new(limits.blobs()),
            BlobBundle::new(limits.blobs()),
        )
        .unwrap();
        assert!(matches!(
            source.apply_durable(&changed),
            Err(Error::Refused(Refusal::Signed))
        ));
    }

    #[test]
    fn joined_move_patches_commute_and_overlap_is_refused() {
        let source = Snapshot::from_bytes(fixture("45543.ppt")).unwrap();
        let moved = |from: usize, destination: usize| {
            let mut edit = source.edit().unwrap();
            edit.move_slide(Position::new(from), Position::new(destination))
                .unwrap();
            edit.commit().unwrap().patch
        };
        let first = moved(0, 1);
        let second = moved(4, 5);
        let overlapping = moved(1, 2);
        let joined = first.join(&second).unwrap();
        assert_eq!(joined.changes().len(), 2);
        assert_eq!(
            slide_texts(joined.apply(&source).unwrap().bytes()),
            slide_texts(second.join(&first).unwrap().apply(&source).unwrap().bytes())
        );
        assert!(matches!(
            joined.join(&overlapping),
            Err(Error::Refused(Refusal::UnsupportedPatch { .. }))
        ));
        let durable = joined.to_durable(patch_limits()).unwrap();
        assert_eq!(
            source.apply_durable(&durable).unwrap(),
            joined.apply(&source).unwrap()
        );
    }

    #[test]
    fn refusals_are_typed_and_failed_moves_leave_the_candidate_unchanged() {
        let source = Snapshot::from_bytes(fixture("basic_test_ppt_file.ppt")).unwrap();
        let mut edit = source.edit().unwrap();
        let error = edit
            .move_slide(Position::new(0), Position::new(99))
            .unwrap_err();
        assert!(matches!(
            error,
            Error::Refused(Refusal::SlideNotFound { position }) if position == Position::new(99)
        ));
        assert!(edit.changes().is_empty());
        assert_eq!(edit.commit().unwrap().snapshot(), &source);

        let mut review_bound = source.clone();
        review_bound.has_review_history = true;
        assert!(matches!(
            review_bound.edit().unwrap_err(),
            Error::Refused(Refusal::ReviewHistoryDependency)
        ));
    }

    fn root_text_publication(
        source: &Snapshot,
        target: crate::text_edit::Target,
        replacement: &str,
    ) -> crate::text_edit::RootPublication {
        let text_source =
            crate::text_edit::Snapshot::from_shared_bytes(source.shared_bytes()).unwrap();
        let mut edit = text_source.edit_text(target).unwrap();
        edit.set_text(replacement).unwrap();
        let commit = edit.commit().unwrap();
        assert!(!commit.patch().is_empty());
        commit.into_root_publication().unwrap()
    }

    #[test]
    fn root_text_publication_adoption_matches_complete_reopen() {
        let source = Snapshot::from_bytes(authored_fixture()).unwrap();
        let target = crate::text_edit::Target::new(Position::new(0), Position::new(0));
        let persist_id = source.document.slides()[0].persist_id();
        let adopted = source
            .adopt_text_edit_publication(
                root_text_publication(&source, target, "adopted root text"),
                persist_id,
            )
            .unwrap();
        let reopened = Snapshot::from_bytes(adopted.bytes().to_vec()).unwrap();

        assert_eq!(adopted, reopened);
        assert_eq!(adopted.document_persist_id, source.document_persist_id);
        assert_eq!(adopted.document, source.document);
        assert_eq!(adopted.has_review_history, source.has_review_history);
        assert_eq!(adopted.limits, source.limits);
        crate::font::validate_unrelated_streams(source.bytes(), adopted.bytes()).unwrap();
    }

    #[test]
    fn root_text_publication_adoption_rejects_wrong_lineage_and_persist_identity() {
        let source = Snapshot::from_bytes(authored_fixture()).unwrap();
        let target = crate::text_edit::Target::new(Position::new(0), Position::new(0));
        let expected = source.document.slides()[0].persist_id();
        let wrong_slide = source.document.slides()[1].persist_id();
        let foreign = Snapshot::from_bytes(fixture("basic_test_ppt_file.ppt")).unwrap();

        let error = foreign
            .adopt_text_edit_publication(
                root_text_publication(&source, target, "foreign root text"),
                expected,
            )
            .unwrap_err();
        assert!(matches!(
            error,
            Error::Package(PackageError::Corrupted(message))
                if message.contains("source changed before adoption")
        ));

        let error = source
            .adopt_text_edit_publication(
                root_text_publication(&source, target, "wrong slide text"),
                wrong_slide,
            )
            .unwrap_err();
        assert!(matches!(
            error,
            Error::Package(PackageError::Corrupted(message))
                if message.contains("unexpected slide record")
        ));

        let error = source
            .adopt_text_edit_publication(
                root_text_publication(&source, target, "document record text"),
                source.document_persist_id,
            )
            .unwrap_err();
        assert!(matches!(
            error,
            Error::Package(PackageError::Corrupted(message))
                if message.contains("live document record")
        ));
    }

    #[test]
    fn root_text_publication_with_custom_limits_retains_complete_reopen() {
        let bytes = authored_fixture();
        let mut limits = SourceLimits::default();
        limits.max_records -= 1;
        let source =
            Snapshot::from_bytes_with_options(bytes, OpenOptions::new().source_limits(limits))
                .unwrap();
        let target = crate::text_edit::Target::new(Position::new(0), Position::new(0));
        let mut edit = source.edit().unwrap();
        edit.set_shape_text(target, "custom limits root text")
            .unwrap();
        let committed = edit.commit().unwrap().snapshot().clone();
        let reopened = Snapshot::from_bytes_with_options(
            committed.bytes().to_vec(),
            OpenOptions::new().source_limits(limits),
        )
        .unwrap();

        assert_eq!(committed, reopened);
        assert_eq!(committed.limits(), limits);
    }

    #[test]
    fn root_transaction_supports_consecutive_adopted_text_publications() {
        let source = Snapshot::from_bytes(authored_fixture()).unwrap();
        let first = crate::text_edit::Target::new(Position::new(0), Position::new(0));
        let second = crate::text_edit::Target::new(Position::new(1), Position::new(0));
        let mut edit = source.edit().unwrap();
        edit.set_shape_text(first, "first adopted text").unwrap();
        edit.set_shape_text(second, "second adopted text").unwrap();
        let commit = edit.commit().unwrap();
        let reopened = Snapshot::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();

        assert_eq!(commit.snapshot(), &reopened);
        assert_eq!(commit.patch().shape_text_changes().len(), 2);
        assert!(slide_texts(reopened.bytes())[0].contains("first adopted text"));
        assert!(slide_texts(reopened.bytes())[1].contains("second adopted text"));
    }

    #[test]
    fn shape_text_batch_rewrites_same_and_cross_slide_records_once_and_reverses() {
        let source = Snapshot::from_bytes(authored_batch_fixture()).unwrap();
        let first = crate::text_edit::Target::new(Position::new(0), Position::new(0));
        let second = crate::text_edit::Target::new(Position::new(0), Position::new(1));
        let third = crate::text_edit::Target::new(Position::new(1), Position::new(0));
        let replacements = [
            ShapeTextReplacement::new(first, "first batch replacement"),
            ShapeTextReplacement::new(second, "second batch replacement is longer"),
            ShapeTextReplacement::new(third, "third batch replacement"),
        ];
        assert_eq!(replacements[0].target(), first);
        assert_eq!(replacements[0].text(), "first batch replacement");

        let mut edit = source.edit().unwrap();
        assert_eq!(edit.set_shape_texts(&replacements).unwrap(), 3);
        assert_eq!(
            edit.shape_text_changes()
                .iter()
                .map(ShapeTextChange::target)
                .collect::<Vec<_>>(),
            [first, second, third]
        );
        let commit = edit.commit().unwrap();
        assert_eq!(
            shape_text(commit.snapshot().bytes(), first),
            "first batch replacement"
        );
        assert_eq!(
            shape_text(commit.snapshot().bytes(), second),
            "second batch replacement is longer"
        );
        assert_eq!(
            shape_text(commit.snapshot().bytes(), third),
            "third batch replacement"
        );
        crate::font::validate_unrelated_streams(source.bytes(), commit.snapshot().bytes()).unwrap();
        let fully_reopened = Snapshot::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
        assert_eq!(commit.snapshot(), &fully_reopened);
        assert_eq!(commit.snapshot().document, source.document);
        assert_eq!(commit.patch().apply(&source).unwrap(), *commit.snapshot());
        assert_eq!(
            commit.patch().inverse().apply(commit.snapshot()).unwrap(),
            source
        );

        let foreign = Snapshot::from_bytes(authored_fixture()).unwrap();
        assert!(commit.patch().apply(&foreign).is_err());
        let durable = commit.patch().to_durable(patch_limits()).unwrap();
        let first_json = durable.to_deterministic_json().unwrap();
        let second_json = durable.to_deterministic_json().unwrap();
        assert_eq!(first_json, second_json);
        let applied = source.apply_durable(&durable).unwrap();
        assert_eq!(
            shape_text(applied.bytes(), first),
            "first batch replacement"
        );
        assert_eq!(
            shape_text(applied.bytes(), second),
            "second batch replacement is longer"
        );
        assert_eq!(
            shape_text(applied.bytes(), third),
            "third batch replacement"
        );
        let restored = applied.apply_durable(&durable.inverse()).unwrap();
        assert_eq!(slide_texts(restored.bytes()), slide_texts(source.bytes()));

        let mut reverse = source.edit().unwrap();
        reverse
            .set_shape_texts(&[
                replacements[2].clone(),
                replacements[1].clone(),
                replacements[0].clone(),
            ])
            .unwrap();
        let reverse_commit = reverse.commit().unwrap();
        assert_eq!(reverse_commit.snapshot().bytes(), commit.snapshot().bytes());
    }

    #[test]
    fn shape_text_batch_refusals_are_atomic_and_duplicate_is_deterministic() {
        let source = Snapshot::from_bytes(authored_batch_fixture()).unwrap();
        let valid = crate::text_edit::Target::new(Position::new(0), Position::new(0));
        let missing = crate::text_edit::Target::new(Position::new(1), Position::new(99));
        let mut edit = source.edit().unwrap();
        let source_ptr = edit.working.bytes().as_ptr();

        let error = edit
            .set_shape_texts(&[
                ShapeTextReplacement::new(valid, "valid but unpublished"),
                ShapeTextReplacement::new(valid, "duplicate"),
            ])
            .unwrap_err();
        assert!(matches!(
            error,
            Error::Refused(Refusal::DuplicateShapeTextTarget { target }) if target == valid
        ));
        assert_eq!(edit.working.bytes().as_ptr(), source_ptr);
        assert!(edit.shape_text_changes().is_empty());

        let error = edit
            .set_shape_texts(&[
                ShapeTextReplacement::new(valid, "valid but unpublished"),
                ShapeTextReplacement::new(missing, "missing"),
            ])
            .unwrap_err();
        assert!(matches!(
            error,
            Error::Text(error)
                if matches!(*error, crate::text_edit::Error::Refused(crate::text_edit::Refusal::ShapeNotFound))
        ));
        assert_eq!(edit.working.bytes().as_ptr(), source_ptr);
        assert!(edit.shape_text_changes().is_empty());

        let error = edit
            .set_shape_texts(&[
                ShapeTextReplacement::new(valid, "valid but unpublished"),
                ShapeTextReplacement::new(
                    crate::text_edit::Target::new(Position::new(0), Position::new(1)),
                    "invalid\0encoding",
                ),
            ])
            .unwrap_err();
        assert!(matches!(
            error,
            Error::Text(error)
                if matches!(*error, crate::text_edit::Error::Refused(crate::text_edit::Refusal::IncompatibleEncoding))
        ));
        assert_eq!(edit.working.bytes().as_ptr(), source_ptr);
        assert!(edit.shape_text_changes().is_empty());
        assert_eq!(edit.commit().unwrap().snapshot(), &source);
    }

    #[test]
    fn shape_text_batch_enforces_limit_before_resolution_and_filters_noops() {
        let source = Snapshot::from_bytes(authored_batch_fixture()).unwrap();
        let target = crate::text_edit::Target::new(Position::new(0), Position::new(0));
        let over_limit = (0..=MAX_SHAPE_TEXT_REPLACEMENTS)
            .map(|_| {
                ShapeTextReplacement::new(
                    crate::text_edit::Target::new(Position::new(usize::MAX), Position::new(0)),
                    "never resolved",
                )
            })
            .collect::<Vec<_>>();
        let mut edit = source.edit().unwrap();
        let error = edit.set_shape_texts(&over_limit).unwrap_err();
        assert!(matches!(
            error,
            Error::Package(PackageError::ResourceLimit(message))
                if message.contains("257") && message.contains("256")
        ));

        let exact_limit = (0..MAX_SHAPE_TEXT_REPLACEMENTS)
            .map(|_| ShapeTextReplacement::new(target, "first shape"))
            .collect::<Vec<_>>();
        assert!(matches!(
            edit.set_shape_texts(&exact_limit),
            Err(Error::Refused(Refusal::DuplicateShapeTextTarget { .. }))
        ));

        let before_ptr = edit.working.bytes().as_ptr();
        assert_eq!(
            edit.set_shape_texts(&[ShapeTextReplacement::new(target, "first shape")])
                .unwrap(),
            0
        );
        assert_eq!(edit.working.bytes().as_ptr(), before_ptr);
        assert!(edit.shape_text_changes().is_empty());

        let second = crate::text_edit::Target::new(Position::new(0), Position::new(1));
        assert_eq!(
            edit.set_shape_texts(&[
                ShapeTextReplacement::new(target, "first shape"),
                ShapeTextReplacement::new(second, "changed second shape"),
            ])
            .unwrap(),
            1
        );
        assert_eq!(edit.shape_text_changes().len(), 1);
        assert_eq!(edit.shape_text_changes()[0].target(), second);

        let exact_source =
            Snapshot::from_bytes(authored_many_shapes_fixture(MAX_SHAPE_TEXT_REPLACEMENTS))
                .unwrap();
        let exact_limit = (0..MAX_SHAPE_TEXT_REPLACEMENTS)
            .map(|shape| {
                ShapeTextReplacement::new(
                    crate::text_edit::Target::new(Position::new(0), Position::new(shape)),
                    format!("shape {shape}"),
                )
            })
            .collect::<Vec<_>>();
        let mut exact_edit = exact_source.edit().unwrap();
        let source_ptr = exact_edit.working.bytes().as_ptr();
        assert_eq!(exact_edit.set_shape_texts(&exact_limit).unwrap(), 0);
        assert_eq!(exact_edit.working.bytes().as_ptr(), source_ptr);
        assert_eq!(exact_edit.commit().unwrap().snapshot(), &exact_source);
    }

    #[test]
    fn shape_text_batch_late_dependency_refusal_leaves_transaction_unchanged() {
        let safe = crate::text_edit::Target::new(Position::new(0), Position::new(0));
        let dependency = crate::text_edit::Target::new(Position::new(0), Position::new(1));
        let bytes = crate::text_edit::add_unmodeled_text_dependency_for_test(
            &authored_batch_fixture(),
            dependency,
        )
        .unwrap();
        let source = Snapshot::from_bytes(bytes).unwrap();
        assert!(!crate::text_edit::shape_text_can_resize(source.bytes(), dependency).unwrap());
        let dependency_before = shape_text(source.bytes(), dependency);
        let mut edit = source.edit().unwrap();
        let source_ptr = edit.working.bytes().as_ptr();
        let error = edit
            .set_shape_texts(&[
                ShapeTextReplacement::new(safe, "valid but unpublished"),
                ShapeTextReplacement::new(dependency, format!("{dependency_before} length change")),
            ])
            .unwrap_err();
        assert!(matches!(
            error,
            Error::Text(error)
                if matches!(*error, crate::text_edit::Error::Refused(crate::text_edit::Refusal::DependencyClosure))
        ));
        assert_eq!(edit.working.bytes().as_ptr(), source_ptr);
        assert!(edit.shape_text_changes().is_empty());
        assert_eq!(edit.commit().unwrap().snapshot(), &source);
    }

    #[test]
    fn shape_text_batch_oversized_late_replacement_is_atomic() {
        let bytes = authored_batch_fixture();
        let limits = SourceLimits {
            max_package_bytes: bytes.len(),
            ..Default::default()
        };
        let source =
            Snapshot::from_bytes_with_options(bytes, OpenOptions::new().source_limits(limits))
                .unwrap();
        let first = crate::text_edit::Target::new(Position::new(0), Position::new(0));
        let second = crate::text_edit::Target::new(Position::new(0), Position::new(1));
        let oversized = "x".repeat(limits.max_package_bytes);
        let mut edit = source.edit().unwrap();
        let source_ptr = edit.working.bytes().as_ptr();
        let error = edit
            .set_shape_texts(&[
                ShapeTextReplacement::new(first, "valid but unpublished"),
                ShapeTextReplacement::new(second, oversized),
            ])
            .unwrap_err();
        assert!(matches!(
            error,
            Error::Text(error)
                if matches!(
                    error.as_ref(),
                    crate::text_edit::Error::Package(PackageError::ResourceLimit(message))
                        if message.contains("shape-text batch retained bytes")
                )
        ));
        assert_eq!(edit.working.bytes().as_ptr(), source_ptr);
        assert!(edit.shape_text_changes().is_empty());
        assert_eq!(edit.commit().unwrap().snapshot(), &source);
    }

    #[test]
    fn shape_text_batch_composes_with_order_conflicts_and_custom_limits() {
        let bytes = authored_batch_fixture();
        let source = Snapshot::from_bytes(bytes.clone()).unwrap();
        let projected_target = crate::text_edit::Target::new(Position::new(1), Position::new(0));
        let mut edit = source.edit().unwrap();
        edit.move_slide(Position::new(0), Position::new(1)).unwrap();
        assert_eq!(
            edit.set_shape_texts(&[ShapeTextReplacement::new(
                projected_target,
                "moved first shape",
            )])
            .unwrap(),
            1
        );
        let commit = edit.commit().unwrap();
        assert_eq!(
            shape_text(commit.snapshot().bytes(), projected_target),
            "moved first shape"
        );

        let original_target = crate::text_edit::Target::new(Position::new(0), Position::new(0));
        let mut competing = source.edit().unwrap();
        competing
            .set_shape_text(original_target, "competing first shape")
            .unwrap();
        let competing = competing.commit().unwrap();
        assert!(
            source
                .plan_three_way(commit.patch(), competing.patch())
                .unwrap()
                .conflicts()
                .iter()
                .any(|conflict| conflict.target() == "slide:0/shape:0/text")
        );

        let mut limits = SourceLimits::default();
        limits.max_records -= 1;
        let limited =
            Snapshot::from_bytes_with_options(bytes, OpenOptions::new().source_limits(limits))
                .unwrap();
        let mut limited_edit = limited.edit().unwrap();
        limited_edit
            .set_shape_texts(&[
                ShapeTextReplacement::new(original_target, "limited first shape"),
                ShapeTextReplacement::new(
                    crate::text_edit::Target::new(Position::new(1), Position::new(0)),
                    "limited third shape",
                ),
            ])
            .unwrap();
        let limited_commit = limited_edit.commit().unwrap();
        let fully_reopened = Snapshot::from_bytes_with_options(
            limited_commit.snapshot().bytes().to_vec(),
            OpenOptions::new().source_limits(limits),
        )
        .unwrap();
        assert_eq!(limited_commit.snapshot(), &fully_reopened);
        assert_eq!(limited_commit.snapshot().limits(), limits);
    }

    #[test]
    fn root_transaction_composes_shape_text_and_order_with_three_way_planning() {
        let source = Snapshot::from_bytes(authored_fixture()).unwrap();
        let target = crate::text_edit::Target::new(Position::new(0), Position::new(0));
        let mut edit = source.edit().unwrap();
        edit.set_shape_text(target, "root transaction text")
            .unwrap();
        let anchor = crate::Anchor::small(20, 30, 300, 200).unwrap();
        edit.set_shape_anchor(target, anchor).unwrap();
        edit.move_slide(Position::new(0), Position::new(1)).unwrap();
        let commit = edit.commit().unwrap();
        let text = slide_texts(commit.snapshot().bytes());
        assert!(text[1].contains("root transaction text"));
        assert_eq!(commit.patch().shape_text_changes().len(), 1);
        assert_eq!(commit.patch().shape_anchor_changes().len(), 1);
        assert_eq!(
            crate::text_edit::inspect_shape_anchor(
                commit.snapshot().bytes(),
                crate::text_edit::Target::new(Position::new(1), Position::new(0)),
            )
            .unwrap(),
            anchor
        );

        let durable = commit.patch().to_durable(patch_limits()).unwrap();
        let applied = source.apply_durable(&durable).unwrap();
        assert_eq!(slide_texts(applied.bytes()), text);
        let restored = applied.apply_durable(&durable.inverse()).unwrap();
        assert_eq!(slide_texts(restored.bytes()), slide_texts(source.bytes()));
        let inverse_patch = commit.patch().inverse().to_durable(patch_limits()).unwrap();
        let inverse_restored = commit.snapshot().apply_durable(&inverse_patch).unwrap();
        assert_eq!(
            slide_texts(inverse_restored.bytes()),
            slide_texts(source.bytes())
        );

        let mut same_target = source.edit().unwrap();
        same_target
            .set_shape_text(target, "competing text")
            .unwrap();
        let competing = same_target.commit().unwrap();
        let conflicts = source
            .plan_three_way(commit.patch(), competing.patch())
            .unwrap();
        assert!(!conflicts.is_clean());
        assert_eq!(conflicts.conflicts().len(), 1);

        let mut disjoint_edit = source.edit().unwrap();
        disjoint_edit
            .set_shape_text(
                crate::text_edit::Target::new(Position::new(1), Position::new(0)),
                "disjoint text",
            )
            .unwrap();
        let disjoint = disjoint_edit.commit().unwrap();
        assert!(
            source
                .plan_three_way(competing.patch(), disjoint.patch())
                .unwrap()
                .is_clean()
        );
    }
    #[test]
    fn real_fixture_remove_is_reversible_on_the_semantic_wire() {
        let source = Snapshot::from_bytes(fixture("basic_test_ppt_file.ppt")).unwrap();
        let original = slide_texts(source.bytes());
        let mut edit = source.edit().unwrap();
        edit.remove_slide(Position::new(0)).unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(commit.snapshot().slide_count(), 1);
        assert_eq!(
            slide_texts(commit.snapshot().bytes()),
            [original[1].clone()]
        );

        let durable = commit.patch().to_durable(transfer_patch_limits()).unwrap();
        let wire = durable.to_deterministic_json().unwrap();
        let decoded =
            litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
                &wire,
                transfer_patch_limits(),
            )
            .unwrap();
        let applied = source.apply_durable(&decoded).unwrap();
        assert_eq!(applied.slide_count(), 1);
        let restored = applied.apply_durable(&decoded.inverse()).unwrap();
        assert_eq!(restored.slide_count(), 2);
        assert_eq!(slide_texts(restored.bytes()), original);
        assert!(
            Snapshot::from_bytes(fixture("45543.ppt"))
                .unwrap()
                .apply_durable(&decoded)
                .is_err()
        );
    }

    #[test]
    fn real_fixture_shape_anchor_reopens_reverses_and_detects_conflicts() {
        let source = Snapshot::from_bytes(fixture("basic_test_ppt_file.ppt")).unwrap();
        let target = anchor_target(&source);
        let before = crate::text_edit::inspect_shape_anchor(source.bytes(), target).unwrap();
        let replacement = crate::Anchor::small(40, 50, 440, 350).unwrap();
        assert_ne!(before, replacement);

        let mut edit = source.edit().unwrap();
        edit.set_shape_anchor(target, replacement).unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(
            crate::text_edit::inspect_shape_anchor(commit.snapshot().bytes(), target).unwrap(),
            replacement
        );
        let mut reopened = Reader::from_reader(Cursor::new(commit.snapshot().bytes())).unwrap();
        assert_eq!(
            reopened.presentation().unwrap().slide_count(),
            source.slide_count()
        );

        let durable = commit.patch().to_durable(patch_limits()).unwrap();
        let applied = source.apply_durable(&durable).unwrap();
        let restored = applied.apply_durable(&durable.inverse()).unwrap();
        assert_eq!(
            crate::text_edit::inspect_shape_anchor(restored.bytes(), target).unwrap(),
            before
        );

        let mut competing_edit = source.edit().unwrap();
        competing_edit
            .set_shape_anchor(target, crate::Anchor::small(60, 70, 460, 370).unwrap())
            .unwrap();
        let competing = competing_edit.commit().unwrap();
        assert!(
            !source
                .plan_three_way(commit.patch(), competing.patch())
                .unwrap()
                .is_clean()
        );
        assert!(competing.snapshot().apply_durable(&durable).is_err());
    }

    #[test]
    fn table_shape_geometry_uses_the_same_checked_root_operation() {
        let source = Snapshot::from_bytes(authored_table_fixture()).unwrap();
        let target = anchor_target(&source);
        let replacement = crate::Anchor::small(80, 90, 880, 590).unwrap();
        let mut edit = source.edit().unwrap();
        edit.set_shape_anchor(target, replacement).unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(commit.snapshot().shape_anchor(target).unwrap(), replacement);
        let durable = commit.patch().to_durable(patch_limits()).unwrap();
        assert_eq!(
            source
                .apply_durable(&durable)
                .unwrap()
                .shape_anchor(target)
                .unwrap(),
            replacement
        );
    }

    #[test]
    fn real_external_media_owner_reopens_and_round_trips_durable_inverse() {
        let (source, id) = real_media_fixture();
        let before = source.external_media_playback(id).unwrap();
        let replacement = crate::external_media::Playback::new(
            !before.loop_playback,
            before.rewind_after_playing,
            before.narration,
        );
        let mut edit = source.edit().unwrap();
        edit.set_external_media_playback(id, replacement).unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(
            commit.snapshot().external_media_playback(id).unwrap(),
            replacement
        );
        assert_eq!(commit.patch().external_media_playback_changes().len(), 1);
        assert_eq!(commit.patch().apply(&source).unwrap(), *commit.snapshot());
        assert_eq!(
            commit.patch().inverse().apply(commit.snapshot()).unwrap(),
            source
        );
        let mut reopened = Reader::from_reader(Cursor::new(commit.snapshot().bytes())).unwrap();
        assert_eq!(
            reopened.presentation().unwrap().slide_count(),
            commit.snapshot().slide_count()
        );

        let durable = commit.patch().to_durable(patch_limits()).unwrap();
        let wire = durable.to_deterministic_json().unwrap();
        let decoded =
            litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
                &wire,
                patch_limits(),
            )
            .unwrap();
        let applied = source.apply_durable(&decoded).unwrap();
        assert_eq!(applied.external_media_playback(id).unwrap(), replacement);
        let restored = applied.apply_durable(&decoded.inverse()).unwrap();
        assert_eq!(restored.external_media_playback(id).unwrap(), before);

        let mut competing_edit = source.edit().unwrap();
        competing_edit
            .set_external_media_playback(
                id,
                crate::external_media::Playback::new(
                    before.loop_playback,
                    before.rewind_after_playing,
                    !before.narration,
                ),
            )
            .unwrap();
        let competing = competing_edit.commit().unwrap();
        assert!(
            !source
                .plan_three_way(commit.patch(), competing.patch())
                .unwrap()
                .is_clean()
        );
    }

    #[test]
    fn live_document_owners_can_be_staged_in_either_order() {
        let (source, id) = real_media_fixture();
        let before = source.external_media_playback(id).unwrap();
        let replacement = crate::external_media::Playback::new(
            !before.loop_playback,
            before.rewind_after_playing,
            before.narration,
        );

        let mut structural_first = source.edit().unwrap();
        structural_first.remove_slide(Position::new(0)).unwrap();
        structural_first
            .set_external_media_playback(id, replacement)
            .unwrap();
        assert_eq!(structural_first.slide_count(), source.slide_count() - 1);
        assert_eq!(structural_first.external_media_playback_changes().len(), 1);
        let structural_first_commit = structural_first.commit().unwrap();
        assert_eq!(
            structural_first_commit
                .snapshot()
                .external_media_playback(id)
                .unwrap(),
            replacement
        );
        assert_eq!(
            structural_first_commit.snapshot().slide_count(),
            source.slide_count() - 1
        );
        assert_eq!(
            structural_first_commit
                .patch()
                .inverse()
                .apply(structural_first_commit.snapshot())
                .unwrap(),
            source
        );

        let mut media_first = source.edit().unwrap();
        media_first
            .set_external_media_playback(id, replacement)
            .unwrap();
        media_first.remove_slide(Position::new(0)).unwrap();
        assert!(media_first.changes().is_empty());
        assert_eq!(media_first.slide_count(), source.slide_count() - 1);
        assert_eq!(media_first.external_media_playback_changes().len(), 1);
        let media_first_commit = media_first.commit().unwrap();
        assert_eq!(
            media_first_commit.snapshot().slide_count(),
            source.slide_count() - 1
        );
        let mut reopened =
            Reader::from_reader(Cursor::new(media_first_commit.snapshot().bytes())).unwrap();
        assert_eq!(
            reopened.presentation().unwrap().slide_count(),
            source.slide_count() - 1
        );
        let durable = media_first_commit
            .patch()
            .to_durable(transfer_patch_limits())
            .unwrap();
        let replayed = source.apply_durable(&durable).unwrap();
        assert_eq!(replayed.slide_count(), source.slide_count() - 1);
        assert_eq!(replayed.external_media_playback(id).unwrap(), replacement);
    }

    #[test]
    fn transfer_planning_reuses_bounded_common_dependencies() {
        let base = Snapshot::from_bytes(authored_fixture()).unwrap();
        let transferred_anchor = crate::Anchor::small(25, 35, 325, 235).unwrap();
        let mut format = base.edit().unwrap();
        format
            .set_shape_anchor(
                crate::text_edit::Target::new(Position::new(0), Position::new(0)),
                transferred_anchor,
            )
            .unwrap();
        let donor = format.commit().unwrap().snapshot().clone();
        let mut remove = donor.edit().unwrap();
        remove.remove_slide(Position::new(0)).unwrap();
        let receiver = remove.commit().unwrap().snapshot().clone();
        let plan = receiver
            .plan_transfer_from(&donor, Position::new(0))
            .unwrap();
        assert!(
            plan.reused_dependencies()
                .contains(&TransferDependency::DrawingGroup)
        );
        let mut edit = receiver.edit().unwrap();
        edit.insert_transfer(Position::new(0), &plan).unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(commit.snapshot().slide_count(), donor.slide_count());
        assert_eq!(
            slide_texts(commit.snapshot().bytes()),
            slide_texts(donor.bytes())
        );
        assert_eq!(
            commit
                .snapshot()
                .shape_anchor(crate::text_edit::Target::new(
                    Position::new(0),
                    Position::new(0),
                ))
                .unwrap(),
            transferred_anchor
        );

        let durable = commit.patch().to_durable(transfer_patch_limits()).unwrap();
        assert_eq!(
            receiver.apply_durable(&durable).unwrap().slide_count(),
            donor.slide_count()
        );

        let drawing = Snapshot::from_bytes(authored_fixture()).unwrap();
        assert!(matches!(
            drawing.plan_transfer_from(&drawing, Position::new(0)),
            Err(Error::Refused(Refusal::UnsupportedSlideDependency {
                dependency: TransferDependency::DrawingGroup
            }))
        ));

        let comments = Snapshot::from_bytes(fixture("WithComments.ppt")).unwrap();
        let mut remove_comment_slide = comments.edit().unwrap();
        remove_comment_slide.remove_slide(Position::new(0)).unwrap();
        let comments_receiver = remove_comment_slide.commit().unwrap().snapshot().clone();
        let comments_plan = comments_receiver
            .plan_transfer_from(&comments, Position::new(0))
            .unwrap();
        assert!(
            comments_plan
                .reused_dependencies()
                .contains(&TransferDependency::CommentAuthorCatalog)
        );
    }

    #[test]
    fn ordinary_drawing_transfer_remaps_into_target_owned_orphan_ids() {
        let donor = Snapshot::from_bytes(authored_fixture()).unwrap();
        let mut remove = donor.edit().unwrap();
        remove.remove_slide(Position::new(1)).unwrap();
        let receiver = remove.commit().unwrap().snapshot().clone();
        let plan = receiver
            .plan_transfer_from(&donor, Position::new(0))
            .unwrap();
        assert!(plan.relationship_remaps().iter().any(|remap| {
            remap.dependency() == TransferDependency::DrawingGroup
                && remap.source_id() != remap.target_id()
        }));

        let mut edit = receiver.edit().unwrap();
        edit.insert_transfer(Position::new(1), &plan).unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(commit.snapshot().slide_count(), 2);
        assert_eq!(
            slide_texts(commit.snapshot().bytes())[0],
            slide_texts(commit.snapshot().bytes())[1]
        );
        let mut reopened = Reader::from_reader(Cursor::new(commit.snapshot().bytes())).unwrap();
        assert_eq!(reopened.presentation().unwrap().slide_count(), 2);

        let durable = commit.patch().to_durable(transfer_patch_limits()).unwrap();
        let applied = receiver.apply_durable(&durable).unwrap();
        assert_eq!(applied.slide_count(), 2);
        let restored = applied.apply_durable(&durable.inverse()).unwrap();
        assert_eq!(restored.slide_count(), 1);
        assert_eq!(
            commit.patch().inverse().apply(commit.snapshot()).unwrap(),
            receiver
        );

        let mut left_edit = receiver.edit().unwrap();
        left_edit.insert_transfer(Position::new(0), &plan).unwrap();
        let left_commit = left_edit.commit().unwrap();
        let mut right_edit = receiver.edit().unwrap();
        right_edit.insert_transfer(Position::new(1), &plan).unwrap();
        let right_commit = right_edit.commit().unwrap();
        let merge = receiver
            .plan_three_way(left_commit.patch(), right_commit.patch())
            .unwrap();
        assert!(
            merge
                .conflicts()
                .iter()
                .any(|conflict| conflict.target().starts_with("drawing/shape-id:"))
        );
    }

    #[test]
    fn novel_picture_transfer_appends_bstore_and_pictures_stream_atomically() {
        let png = vec![0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4];
        let mut donor_writer = crate::writer::Builder::new();
        let donor_slide = donor_writer.add_slide().unwrap();
        donor_writer
            .add_picture_as(
                donor_slide,
                20,
                30,
                200,
                120,
                png.clone(),
                crate::writer::PictureKind::Png,
            )
            .unwrap();
        let mut donor_bytes = Cursor::new(Vec::new());
        crate::raw::encode(&mut donor_writer, &mut donor_bytes).unwrap();
        let donor = Snapshot::from_bytes(donor_bytes.into_inner()).unwrap();

        let mut target_writer = crate::writer::Builder::new();
        let live = target_writer.add_slide().unwrap();
        target_writer
            .add_textbox(live, 10, 10, 100, 30, "receiver")
            .unwrap();
        let orphan = target_writer.add_slide().unwrap();
        for offset in [0, 30, 60, 90] {
            target_writer
                .add_rectangle(orphan, 10 + offset, 20, 20, 20)
                .unwrap();
        }
        let mut target_bytes = Cursor::new(Vec::new());
        crate::raw::encode(&mut target_writer, &mut target_bytes).unwrap();
        let target = Snapshot::from_bytes(target_bytes.into_inner()).unwrap();
        let mut remove = target.edit().unwrap();
        remove.remove_slide(Position::new(1)).unwrap();
        let receiver = remove.commit().unwrap().snapshot().clone();
        assert!(picture_inventory(&receiver).unwrap().resources.is_empty());
        assert!(picture_stream(&receiver).unwrap().is_empty());

        let plan = receiver
            .plan_transfer_from(&donor, Position::new(0))
            .unwrap();
        assert_eq!(plan.payload.pictures.len(), 1);
        assert_eq!(plan.payload.pictures[0].target_id, 1);
        let mut edit = receiver.edit().unwrap();
        edit.insert_transfer(Position::new(1), &plan).unwrap();
        let commit = edit.commit().unwrap();
        let inventory = picture_inventory(commit.snapshot()).unwrap();
        assert_eq!(inventory.slots, 1);
        assert_eq!(inventory.resources[&1].data, png);
        assert!(!picture_stream(commit.snapshot()).unwrap().is_empty());

        let durable = commit.patch().to_durable(transfer_patch_limits()).unwrap();
        let applied = receiver.apply_durable(&durable).unwrap();
        assert_eq!(picture_inventory(&applied).unwrap().resources[&1].data, png);
        assert_eq!(applied.slide_count(), 2);
        assert_eq!(
            commit.patch().inverse().apply(commit.snapshot()).unwrap(),
            receiver
        );
    }

    #[test]
    fn relationship_atom_remapping_is_fixed_width_and_fail_closed() {
        let mut atom = crate::InteractiveInfoAtom {
            sound_id: 7,
            hyperlink_id: 11,
            action: crate::InteractionAction::Hyperlink,
            ole_verb: 0,
            jump: crate::InteractionJump::None,
            animated: false,
            stop_sound: false,
            custom_show_return: false,
            visited: false,
            link_target: crate::InteractionLinkTarget::Url,
            unused: [0; 3],
        }
        .to_payload();
        rewrite_interactive_atom(
            &mut atom,
            &BTreeMap::from([(11, 101)]),
            &BTreeMap::from([(7, 107)]),
        )
        .unwrap();
        let parsed = crate::InteractiveInfoAtom::parse_payload(&atom).unwrap();
        assert_eq!(parsed.hyperlink_id, 101);
        assert_eq!(parsed.sound_id, 107);
        assert_eq!(atom.len(), 16);

        let mut active = atom;
        active[8] = 5;
        let mut relationships = SlideRelationships::default();
        scan_interactive_atom(&active, &mut relationships).unwrap();
        assert!(relationships.active_action);
    }

    #[test]
    fn connector_rules_remap_three_shape_ids_without_changing_record_width() {
        let mut payload = Vec::new();
        for value in [1_u32, 0x0401, 0x0402, 0x0403, 7, 9] {
            payload.extend_from_slice(&value.to_le_bytes());
        }
        let mut record = Vec::new();
        record.extend_from_slice(&1_u16.to_le_bytes());
        record.extend_from_slice(&0xF012_u16.to_le_bytes());
        record.extend_from_slice(&24_u32.to_le_bytes());
        record.extend_from_slice(&payload);
        let before_len = record.len();

        let mut relationships = SlideRelationships::default();
        scan_officeart_records(&record, &mut relationships).unwrap();
        assert_eq!(
            relationships.connector_shape_ids,
            BTreeSet::from([0x0401, 0x0402, 0x0403])
        );

        rewrite_officeart_records(
            &mut record,
            &BTreeMap::new(),
            &BTreeMap::new(),
            &BTreeMap::new(),
            &BTreeMap::new(),
            &BTreeMap::from([(0x0401, 0x0801), (0x0402, 0x0802), (0x0403, 0x0803)]),
        )
        .unwrap();
        assert_eq!(record.len(), before_len);
        assert_eq!(read_payload_u32(&record[8..], 0).unwrap(), 1);
        assert_eq!(read_payload_u32(&record[8..], 4).unwrap(), 0x0801);
        assert_eq!(read_payload_u32(&record[8..], 8).unwrap(), 0x0802);
        assert_eq!(read_payload_u32(&record[8..], 12).unwrap(), 0x0803);
        assert_eq!(read_payload_u32(&record[8..], 16).unwrap(), 7);
        assert_eq!(read_payload_u32(&record[8..], 20).unwrap(), 9);

        assert!(rewrite_connector_rule(&mut payload, &BTreeMap::from([(0x0401, 0x0801)])).is_err());
    }

    #[test]
    fn officeart_shape_and_blip_reference_families_close_under_remapping() {
        fn push_record(output: &mut Vec<u8>, header: u16, kind: u16, payload: &[u8]) {
            output.extend_from_slice(&header.to_le_bytes());
            output.extend_from_slice(&kind.to_le_bytes());
            output.extend_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
            output.extend_from_slice(payload);
        }

        let mut drawing = Vec::new();
        let mut options = Vec::new();
        options.extend_from_slice(&0x008a_u16.to_le_bytes());
        options.extend_from_slice(&0x0401_u32.to_le_bytes());
        options.extend_from_slice(&0x4104_u16.to_le_bytes());
        options.extend_from_slice(&5_u32.to_le_bytes());
        push_record(&mut drawing, 0x0023, 0xF00B, &options);

        for (kind, shape_id) in [(0xF014_u16, 0x0402_u32), (0xF017, 0x0403)] {
            let mut payload = 7_u32.to_le_bytes().to_vec();
            payload.extend_from_slice(&shape_id.to_le_bytes());
            push_record(&mut drawing, 0, kind, &payload);
        }
        push_record(
            &mut drawing,
            0,
            0xF11D,
            &(0xc000_0000_u32 | 0x0404).to_le_bytes(),
        );
        let mut selection = vec![0; 8];
        selection.extend_from_slice(&0x0405_u32.to_le_bytes());
        selection.extend_from_slice(&0x0406_u32.to_le_bytes());
        push_record(&mut drawing, 0, 0xF119, &selection);

        let mut relationships = SlideRelationships::default();
        scan_officeart_records(&drawing, &mut relationships).unwrap();
        assert_eq!(relationships.blip_ids, BTreeSet::from([5]));
        assert_eq!(
            relationships.other_shape_reference_ids,
            BTreeSet::from([0x0401, 0x0402, 0x0403, 0x0404, 0x0405, 0x0406])
        );

        let drawings = BTreeMap::from([
            (0x0401, 0x0801),
            (0x0402, 0x0802),
            (0x0403, 0x0803),
            (0x0404, 0x0804),
            (0x0405, 0x0805),
            (0x0406, 0x0806),
        ]);
        rewrite_officeart_records(
            &mut drawing,
            &BTreeMap::new(),
            &BTreeMap::new(),
            &BTreeMap::new(),
            &BTreeMap::from([(5, 9)]),
            &drawings,
        )
        .unwrap();

        let mut rewritten = SlideRelationships::default();
        scan_officeart_records(&drawing, &mut rewritten).unwrap();
        assert_eq!(rewritten.blip_ids, BTreeSet::from([9]));
        assert_eq!(
            rewritten.other_shape_reference_ids,
            BTreeSet::from([0x0801, 0x0802, 0x0803, 0x0804, 0x0805, 0x0806])
        );
        let mut deleted = (0xc000_0000_u32 | 0x0404).to_le_bytes();
        rewrite_deleted_shape_rule(&mut deleted, &drawings).unwrap();
        assert_eq!(u32::from_le_bytes(deleted), 0xc000_0000_u32 | 0x0804);
    }

    #[test]
    fn external_shape_and_animation_refusals_are_distinct() {
        let source = Snapshot::from_bytes(fixture("basic_test_ppt_file.ppt")).unwrap();
        let mut linked_shape_property = Vec::new();
        linked_shape_property.extend_from_slice(&0x008a_u16.to_le_bytes());
        linked_shape_property.extend_from_slice(&0x0401_u32.to_le_bytes());
        let mut relationships = SlideRelationships::default();
        scan_linked_shape_references(1, &linked_shape_property, &mut relationships).unwrap();
        assert_eq!(
            relationships.other_shape_reference_ids,
            BTreeSet::from([0x0401])
        );
        linked_shape_property[0..2].copy_from_slice(&0x4001_u16.to_le_bytes());
        assert_eq!(
            option_blip_references(1, &linked_shape_property).unwrap(),
            BTreeSet::from([0x0401])
        );

        let connector = SlideRelationships {
            connector_shape_ids: BTreeSet::from([0x0401]),
            ..SlideRelationships::default()
        };
        assert!(matches!(
            drawing_remaps(&source, &connector),
            Err(Error::Refused(Refusal::UnsupportedSlideDependency {
                dependency: TransferDependency::ConnectorRule
            }))
        ));

        let other_reference = SlideRelationships {
            unsupported_dependencies: BTreeSet::from([
                TransferDependency::OtherOfficeArtShapeReference,
            ]),
            ..SlideRelationships::default()
        };
        assert!(matches!(
            drawing_remaps(&source, &other_reference),
            Err(Error::Refused(Refusal::UnsupportedSlideDependency {
                dependency: TransferDependency::OtherOfficeArtShapeReference
            }))
        ));

        let animation = SlideRelationships {
            unsupported_dependencies: BTreeSet::from([TransferDependency::AnimationBuildGraph]),
            ..SlideRelationships::default()
        };
        assert!(matches!(
            drawing_remaps(&source, &animation),
            Err(Error::Refused(Refusal::UnsupportedSlideDependency {
                dependency: TransferDependency::AnimationBuildGraph
            }))
        ));
    }

    /// The slide container's child record types.
    fn slide_children(snapshot: &Snapshot, position: usize) -> Vec<crate::RecordType> {
        let slide = snapshot.document.slides()[position];
        let bytes = persisted_record(snapshot, slide.persist_id()).unwrap();
        let (root, _) = crate::Record::parse(&bytes, 0, &snapshot.limits).unwrap();
        root.children
            .iter()
            .map(|child| child.record_type)
            .collect()
    }

    /// Record 0290 F4: a slide without slide-show information (here, from
    /// the PPT writer) gains a `SlideShowSlideInfoAtom` after its `SlideAtom`
    /// when it is hidden; showing it is then a flag change, and the exact
    /// patch and its inverse restore both states.
    #[test]
    fn slide_visibility_inserts_a_missing_slide_show_information_atom() {
        let source = Snapshot::from_bytes(authored_fixture()).unwrap();
        let atom = crate::RecordType::SSSlideInfoAtom;
        assert!(!slide_children(&source, 0).contains(&atom));
        assert!(!source.slide_hidden(Position::new(0)).unwrap());

        let mut noop = source.edit().unwrap();
        noop.set_slide_hidden(Position::new(0), false).unwrap();
        assert_eq!(noop.commit().unwrap().snapshot().bytes(), source.bytes());

        let mut edit = source.edit().unwrap();
        edit.set_slide_hidden(Position::new(0), true).unwrap();
        let commit = edit.commit().unwrap();
        let hidden = commit.snapshot();
        assert!(hidden.slide_hidden(Position::new(0)).unwrap());
        assert!(!hidden.slide_hidden(Position::new(1)).unwrap());
        let children = slide_children(hidden, 0);
        assert_eq!(children[0], crate::RecordType::SlideAtom);
        assert_eq!(children[1], atom);
        assert_eq!(children.iter().filter(|kind| **kind == atom).count(), 1);
        assert_eq!(slide_children(hidden, 1), slide_children(&source, 1));
        assert_eq!(slide_texts(hidden.bytes()), slide_texts(source.bytes()));
        // The inserted atom carries no transition, manual advance.
        assert_eq!(
            hidden.slide_advance(Position::new(0)).unwrap(),
            SlideAdvance::on_click()
        );
        assert_eq!(commit.patch().apply(&source).unwrap(), *hidden);
        assert_eq!(commit.patch().inverse().apply(hidden).unwrap(), source);

        let mut shown = hidden.edit().unwrap();
        shown.set_slide_hidden(Position::new(0), false).unwrap();
        let shown = shown.commit().unwrap();
        assert!(!shown.snapshot().slide_hidden(Position::new(0)).unwrap());
        assert_eq!(slide_children(shown.snapshot(), 0), children);
    }

    #[test]
    fn slide_visibility_composes_with_order_and_round_trips_durably() {
        let source = Snapshot::from_bytes(fixture("45543.ppt")).unwrap();
        let original = source.slide_hidden(Position::new(0)).unwrap();
        let slide_id = source.document.slides()[0].slide_id();
        let non_outline_flags = source.document.slides()[0].flags();
        let mut edit = source.edit().unwrap();
        edit.set_slide_hidden(Position::new(0), !original).unwrap();
        edit.move_slide(Position::new(0), Position::new(1)).unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(commit.patch().slide_visibility_changes().len(), 1);
        assert_eq!(
            commit.snapshot().slide_hidden(Position::new(1)).unwrap(),
            !original
        );
        let moved = commit
            .snapshot()
            .document
            .slides()
            .iter()
            .find(|slide| slide.slide_id() == slide_id)
            .unwrap();
        assert_eq!(
            moved.flags(),
            non_outline_flags,
            "set_slide_hidden must preserve SlidePersistAtom.fNonOutlineData"
        );
        assert_eq!(commit.patch().apply(&source).unwrap(), *commit.snapshot());
        assert_eq!(
            commit.patch().inverse().apply(commit.snapshot()).unwrap(),
            source
        );

        let durable = commit.patch().to_durable(patch_limits()).unwrap();
        let applied = source.apply_durable(&durable).unwrap();
        assert_eq!(applied, *commit.snapshot());
        let visibility_restored = applied.apply_durable(&durable.inverse()).unwrap();
        assert_eq!(
            visibility_restored.document.slides(),
            source.document.slides()
        );
        assert_eq!(
            slide_texts(visibility_restored.bytes()),
            slide_texts(source.bytes())
        );
        let mut reopened = Reader::from_reader(Cursor::new(visibility_restored.bytes())).unwrap();
        assert_eq!(
            reopened.presentation().unwrap().slide_count(),
            source.slide_count()
        );

        let mut left_edit = source.edit().unwrap();
        left_edit
            .set_slide_hidden(Position::new(0), !original)
            .unwrap();
        let left_commit = left_edit.commit().unwrap();
        let mut right_edit = source.edit().unwrap();
        right_edit
            .set_slide_hidden(Position::new(0), !original)
            .unwrap();
        let right_commit = right_edit.commit().unwrap();
        let expected_conflict = format!(
            "slide-id:{}/hidden",
            left_commit.patch().slide_visibility_changes()[0].slide_id
        );
        assert_eq!(
            source
                .plan_three_way(left_commit.patch(), right_commit.patch())
                .unwrap()
                .conflicts()[0]
                .target(),
            expected_conflict
        );
    }

    #[test]
    fn genuine_corpus_slide_advance_reopens_reverses_merges_and_undoes() {
        let cases = genuine_advance_cases();
        let producer_files = cases
            .iter()
            .filter_map(|(family, name, _, _)| (*family == "poi-slideshow").then_some(*name))
            .collect::<BTreeSet<_>>();
        assert!(
            producer_files.is_superset(&BTreeSet::from(["45543.ppt", "WithLinks.ppt"])),
            "the two checked-in genuine producer files must expose slide-show information"
        );

        for (family, name, source, position) in &cases {
            let before = source.slide_advance(*position).unwrap();
            let replacement = if before == SlideAdvance::both(1_750).unwrap() {
                SlideAdvance::on_click()
            } else {
                SlideAdvance::both(1_750).unwrap()
            };
            let before_atom = advance_atom_data(source, *position);
            let mut edit = source.edit().unwrap();
            edit.set_slide_advance(*position, replacement).unwrap();
            let commit = edit.commit().unwrap();
            assert_eq!(
                commit.snapshot().slide_advance(*position).unwrap(),
                replacement
            );
            let after_atom = advance_atom_data(commit.snapshot(), *position);
            assert_eq!(&after_atom[4..10], &before_atom[4..10], "{family}/{name}");
            assert_eq!(&after_atom[12..16], &before_atom[12..16], "{family}/{name}");
            let mut reopened = Reader::from_reader(Cursor::new(commit.snapshot().bytes())).unwrap();
            assert_eq!(
                reopened.presentation().unwrap().slide_count(),
                source.slide_count(),
                "{family}/{name}"
            );
        }

        let (_family, _name, source, position) = cases.into_iter().next().unwrap();
        let before = source.slide_advance(position).unwrap();
        let replacement = if before == SlideAdvance::automatic(2_250).unwrap() {
            SlideAdvance::both(2_250).unwrap()
        } else {
            SlideAdvance::automatic(2_250).unwrap()
        };
        let mut edit = source.edit().unwrap();
        edit.set_slide_advance(position, replacement).unwrap();
        let destination = if source.slide_count() > 1 {
            if position.get() == 0 {
                Position::new(1)
            } else {
                Position::new(0)
            }
        } else {
            position
        };
        if destination != position {
            edit.move_slide(position, destination).unwrap();
        }
        let commit = edit.commit().unwrap();
        assert_eq!(commit.patch().slide_advance_changes().len(), 1);
        assert_eq!(
            commit.snapshot().slide_advance(destination).unwrap(),
            replacement
        );
        assert_eq!(commit.patch().apply(&source).unwrap(), *commit.snapshot());
        assert_eq!(
            commit.patch().inverse().apply(commit.snapshot()).unwrap(),
            source
        );

        let durable = commit.patch().to_durable(patch_limits()).unwrap();
        let wire = durable.to_deterministic_json().unwrap();
        let decoded =
            litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
                &wire,
                patch_limits(),
            )
            .unwrap();
        let applied = source.apply_durable(&decoded).unwrap();
        assert_eq!(applied, *commit.snapshot());
        let restored = applied.apply_durable(&decoded.inverse()).unwrap();
        assert_eq!(restored.slide_advance(position).unwrap(), before);
        assert_eq!(slide_texts(restored.bytes()), slide_texts(source.bytes()));

        // Undo and redo are the inverse and the patch applied again.
        let undone = commit.patch().inverse().apply(commit.snapshot()).unwrap();
        assert_eq!(undone, source);
        assert_eq!(commit.patch().apply(&undone).unwrap(), *commit.snapshot());

        let mut competing_edit = source.edit().unwrap();
        competing_edit
            .set_slide_advance(position, SlideAdvance::both(3_500).unwrap())
            .unwrap();
        let competing_commit = competing_edit.commit().unwrap();
        assert!(
            source
                .plan_three_way(commit.patch(), competing_commit.patch())
                .unwrap()
                .conflicts()
                .iter()
                .any(|conflict| conflict.target().ends_with("/advance"))
        );
    }

    #[test]
    fn slide_advance_refuses_missing_owner_and_out_of_range_delay() {
        assert!(SlideAdvance::automatic(MAX_SLIDE_ADVANCE_MS + 1).is_err());
        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/ole/ppt/SampleShow.ppt"),
        )
        .unwrap();
        let source = Snapshot::from_bytes(bytes).unwrap();
        let position = Position::new(0);
        assert!(matches!(
            source.slide_advance(position),
            Err(Error::Refused(Refusal::UnsupportedSlideAdvance { position: refused }))
                if refused == position
        ));
        let mut edit = source.edit().unwrap();
        assert!(matches!(
            edit.set_slide_advance(position, SlideAdvance::on_click()),
            Err(Error::Refused(Refusal::UnsupportedSlideAdvance { position: refused }))
                if refused == position
        ));
        assert!(edit.slide_advance_changes().is_empty());
        assert_eq!(edit.commit().unwrap().snapshot(), &source);
    }

    fn memoized_digest(snapshot: &Snapshot) -> Option<&str> {
        snapshot.artifact.memoized_sha256()
    }

    fn exact_digest(bytes: &[u8]) -> String {
        litchi_core::patch::BlobId::of(bytes).as_hex()
    }

    fn artifact_preconditions(operations: &[litchi_core::patch::PatchOperation]) -> Vec<&str> {
        operations
            .iter()
            .filter_map(|operation| operation.preconditions.get("artifact_sha256"))
            .map(|value| value.as_str().unwrap())
            .collect()
    }

    /// Asserts that durable application stopped at the structural
    /// `artifact_sha256` precondition rather than at a formatting, order or
    /// payload precondition.
    fn assert_structural_artifact_conflict(result: Result<Snapshot>) {
        match result {
            Err(Error::Package(PackageError::InvalidFormat(message))) => assert_eq!(
                message,
                "PPT durable structural patch source artifact does not match"
            ),
            Err(error) => panic!("expected the structural artifact conflict, got: {error}"),
            Ok(_) => panic!("a durable structural patch applied to a sibling artifact"),
        }
    }

    #[test]
    fn commit_defers_artifact_digests_to_durable_serialization() {
        let source = Snapshot::from_bytes(fixture("45543.ppt")).unwrap();
        let mut edit = source.edit().unwrap();
        edit.remove_slide(Position::new(1)).unwrap();
        let commit = edit.commit().unwrap();
        let patch = commit.patch();
        assert_eq!(patch.structural_before, StructuralArtifact::Before);
        assert_eq!(patch.structural_after, StructuralArtifact::After);
        // Commit and exact in-memory application compute no artifact digest.
        assert_eq!(patch.apply(&source).unwrap(), *commit.snapshot());
        assert_eq!(patch.inverse().apply(commit.snapshot()).unwrap(), source);
        for snapshot in [&source, &patch.before, &patch.after, commit.snapshot()] {
            assert_eq!(memoized_digest(snapshot), None);
        }

        let durable = patch.to_durable(transfer_patch_limits()).unwrap();
        let source_digest = exact_digest(source.bytes());
        let target_digest = exact_digest(commit.snapshot().bytes());
        assert_ne!(source_digest, target_digest);
        // The digests were computed from exactly the retained bytes and are
        // shared with every clone: the caller's source and committed snapshot.
        for (snapshot, expected) in [
            (&source, &source_digest),
            (&patch.before, &source_digest),
            (&patch.after, &target_digest),
            (commit.snapshot(), &target_digest),
        ] {
            assert_eq!(memoized_digest(snapshot), Some(expected.as_str()));
        }
        assert_eq!(
            artifact_preconditions(durable.operations()),
            [source_digest.as_str()]
        );
        assert_eq!(
            artifact_preconditions(durable.inverse().operations()),
            [target_digest.as_str()]
        );
        // The inverse patch reuses both memoized digests with swapped roles.
        let inverse = patch.inverse().to_durable(transfer_patch_limits()).unwrap();
        assert_eq!(
            artifact_preconditions(inverse.operations()),
            [target_digest.as_str()]
        );
        assert_eq!(
            artifact_preconditions(inverse.inverse().operations()),
            [source_digest.as_str()]
        );
        // Repeated serialization is deterministic and byte-identical.
        assert_eq!(
            patch
                .to_durable(transfer_patch_limits())
                .unwrap()
                .to_deterministic_json()
                .unwrap(),
            durable.to_deterministic_json().unwrap()
        );

        // A following edit of the committed snapshot starts from its memo.
        let mut next = commit.snapshot().edit().unwrap();
        next.remove_slide(Position::new(1)).unwrap();
        let next = next.commit().unwrap();
        assert_eq!(
            memoized_digest(&next.patch().before),
            Some(target_digest.as_str())
        );
        assert_eq!(memoized_digest(&next.patch().after), None);
        let next_durable = next.patch().to_durable(transfer_patch_limits()).unwrap();
        assert_eq!(
            artifact_preconditions(next_durable.operations()),
            [target_digest.as_str()]
        );
        assert_eq!(
            artifact_preconditions(next_durable.inverse().operations()),
            [exact_digest(next.snapshot().bytes()).as_str()]
        );
    }

    /// Asserts the SHA-256 and length of a commit's forward and inverse
    /// deterministic durable JSON.
    fn assert_durable_wire(
        label: &str,
        commit: &Commit,
        committed: (&str, usize),
        forward: (&str, usize),
        inverse: (&str, usize),
    ) {
        let digest = |bytes: &[u8]| (exact_digest(bytes), bytes.len());
        let wire = |patch: &Patch| {
            patch
                .to_durable(transfer_patch_limits())
                .unwrap()
                .to_deterministic_json()
                .unwrap()
        };
        let expected = |(sha, len): (&str, usize)| (sha.to_string(), len);
        assert_eq!(
            digest(commit.snapshot().bytes()),
            expected(committed),
            "{label}: committed artifact"
        );
        assert_eq!(
            digest(&wire(commit.patch())),
            expected(forward),
            "{label}: forward durable JSON"
        );
        assert_eq!(
            digest(&wire(&commit.patch().inverse())),
            expected(inverse),
            "{label}: inverse durable JSON"
        );
    }

    /// The expected values were produced by the implementation that computed
    /// both structural artifact digests eagerly at commit (base `009d515bef`),
    /// using change 0745's probe over the same fixtures, edits and limits.
    /// Deferring the digests must leave every durable byte unchanged.
    #[test]
    fn durable_wire_bytes_match_the_eager_digest_implementation() {
        let deck = || Snapshot::from_bytes(fixture("45543.ppt")).unwrap();

        let source = deck();
        let mut edit = source.edit().unwrap();
        edit.remove_slide(Position::new(1)).unwrap();
        assert_durable_wire(
            "45543 remove slide 1",
            &edit.commit().unwrap(),
            (
                "545ed7e5ba7b8cc0687f8d21ac95aca68fa4d67fd89fd4277ed7cce047bdb859",
                390_144,
            ),
            (
                "dba5682ef8d8f68fa1900c5791dfa88d1330d706df17d551ae749531287af87f",
                8_351,
            ),
            (
                "2f481aa5e19dbb60c3f80c0f77bdfe5da9677eb8e0ca5e1ae6c53f211334a8ba",
                8_351,
            ),
        );

        let source = deck();
        let mut edit = source.edit().unwrap();
        edit.move_slide(Position::new(0), Position::new(source.slide_count() - 1))
            .unwrap();
        assert_durable_wire(
            "45543 move slide 0 to last",
            &edit.commit().unwrap(),
            (
                "119e23514488f04f97c8c0b04a27b85a05553cf572efaa8ce0cfd8260c0df4ce",
                390_144,
            ),
            (
                "649801102f8397b5537f8312c50f03d2ec8dfb0fe048f7246bf1f11cfb696305",
                592,
            ),
            (
                "6a5c47191214e68a2b493bf629359efd969b94cf9437ba04ff4d12891a56e3d8",
                592,
            ),
        );

        let source = deck();
        let hidden = source.slide_hidden(Position::new(0)).unwrap();
        let mut edit = source.edit().unwrap();
        edit.set_slide_hidden(Position::new(0), !hidden).unwrap();
        assert_durable_wire(
            "45543 formatting-only hide",
            &edit.commit().unwrap(),
            (
                "8209e62cf70be791091cca8593490bc0dbf49184234f7d28c29ce8cdbc79c9cc",
                397_312,
            ),
            (
                "8caf2cb4ec56b99ee58a4872ac12c35aeaf0e15bf167453f8242b251de899fca",
                318,
            ),
            (
                "d643a40400d4ff6e5cff7562344b909879355b5937cf1e26e021eee4135574e2",
                318,
            ),
        );

        let source = deck();
        let mut edit = source.edit().unwrap();
        edit.set_slide_hidden(Position::new(0), !hidden).unwrap();
        edit.remove_slide(Position::new(1)).unwrap();
        assert_durable_wire(
            "45543 hide then remove (intermediate artifact)",
            &edit.commit().unwrap(),
            (
                "13b7b21048729ccb068e7ee7cd372060feb4ddedbd93a121d59a75f75841a5e2",
                402_432,
            ),
            (
                "650d815733819b44225ec159001b63bccd90fde5a815f2a361f85fb5d61a412c",
                8_581,
            ),
            (
                "7c54b006f5e2d46bd2b288c14e08fd3e0345e63d1625f14eb1170f6fc9253e35",
                8_581,
            ),
        );

        let source = deck();
        let mut edit = source.edit().unwrap();
        edit.move_slide(Position::new(0), Position::new(1)).unwrap();
        edit.move_slide(Position::new(1), Position::new(0)).unwrap();
        let net_zero = (
            "1328c0225965fbe9fd74eaf240e202ff77d3fba075c212604cd7a287b159bb1a",
            1_092,
        );
        assert_durable_wire(
            "45543 net-zero moves",
            &edit.commit().unwrap(),
            (
                "218aaac542e5f9b567736407f2631defc65797c6ba2a7818f066e2f93bcfacaf",
                385_024,
            ),
            net_zero,
            net_zero,
        );

        let authored = || Snapshot::from_bytes(authored_fixture()).unwrap();
        let target = crate::text_edit::Target::new(Position::new(0), Position::new(0));

        let source = authored();
        let mut edit = source.edit().unwrap();
        edit.set_shape_text(target, "golden replacement").unwrap();
        edit.move_slide(Position::new(0), Position::new(1)).unwrap();
        assert_durable_wire(
            "authored text then move (intermediate artifact)",
            &edit.commit().unwrap(),
            (
                "719cc3c9e888d8f26e050d26c499fe8607f793192af977f337d01b1495af4b7d",
                9_216,
            ),
            (
                "ca74b8104062c56283053b080911a3e4e91dbeb7c175b86c8edae49976f0aad6",
                989,
            ),
            (
                "d3b26915f3a722280b2759d2f8729215fc0a942235c11c7ade9556f62f66674f",
                989,
            ),
        );

        let source = authored();
        let mut edit = source.edit().unwrap();
        edit.set_shape_anchor(target, crate::Anchor::small(25, 35, 325, 235).unwrap())
            .unwrap();
        edit.remove_slide(Position::new(1)).unwrap();
        assert_durable_wire(
            "authored anchor then remove (intermediate artifact)",
            &edit.commit().unwrap(),
            (
                "c024ea35d572a5d83ddb009041c0c26fc949038efba2dbc051a473daa12b3b0d",
                9_216,
            ),
            (
                "c25870460dfc2b979fb643d8fb737f7263f394c1acaff777993cc435b820cac7",
                2_046,
            ),
            (
                "76e51674df051c182a4353dd44a27aafeff53949211b1a555c290918a65ca8fe",
                2_046,
            ),
        );

        let donor = authored();
        let mut remove = donor.edit().unwrap();
        remove.remove_slide(Position::new(1)).unwrap();
        let receiver = remove.commit().unwrap().snapshot().clone();
        let plan = receiver
            .plan_transfer_from(&donor, Position::new(0))
            .unwrap();
        let mut edit = receiver.edit().unwrap();
        edit.insert_transfer(Position::new(1), &plan).unwrap();
        assert_durable_wire(
            "authored transfer insert",
            &edit.commit().unwrap(),
            (
                "fc73e8a5aed23674f7f0f2abb107dd2b2facd010845ed6c68eed7f750470c204",
                9_728,
            ),
            (
                "37cced31a154e508262ab213bc97b1be0a1b1925b1d8a8970b4d681a1b50d366",
                1_543,
            ),
            (
                "5c880f90a56401a573f312b1aca6e2267c6825c4789b7f83a15cb44001d52c00",
                1_543,
            ),
        );
    }

    #[test]
    fn exact_noop_and_formatting_only_commits_compute_no_artifact_digest() {
        let source = Snapshot::from_bytes(fixture("45543.ppt")).unwrap();
        let noop = source.edit().unwrap().commit().unwrap();
        assert!(noop.patch().is_empty());
        assert!(std::ptr::eq(
            source.bytes().as_ptr(),
            noop.snapshot().bytes().as_ptr()
        ));
        assert_eq!(noop.patch().structural_before, StructuralArtifact::After);
        assert_eq!(noop.patch().structural_after, StructuralArtifact::After);
        assert_eq!(memoized_digest(&source), None);
        let noop_durable = noop.patch().to_durable(patch_limits()).unwrap();
        assert!(noop_durable.operations().is_empty());
        assert_eq!(memoized_digest(&source), None);

        let hidden = source.slide_hidden(Position::new(0)).unwrap();
        let mut format = source.edit().unwrap();
        format.set_slide_hidden(Position::new(0), !hidden).unwrap();
        let format = format.commit().unwrap();
        assert!(!format.patch().is_empty());
        assert_eq!(memoized_digest(&format.patch().before), None);
        assert_eq!(memoized_digest(&format.patch().after), None);
        // Formatting operations bind semantic values, not artifact digests,
        // so durable serialization of a formatting-only patch needs none.
        let durable = format.patch().to_durable(patch_limits()).unwrap();
        assert!(artifact_preconditions(durable.operations()).is_empty());
        assert_eq!(memoized_digest(&format.patch().before), None);
        assert_eq!(memoized_digest(&format.patch().after), None);
        assert_eq!(
            source.apply_durable(&durable).unwrap().bytes(),
            format.snapshot().bytes()
        );

        // Net-zero staged moves serialize against the retained target.
        let mut round_trip = source.edit().unwrap();
        round_trip
            .move_slide(Position::new(0), Position::new(1))
            .unwrap();
        round_trip
            .move_slide(Position::new(1), Position::new(0))
            .unwrap();
        let round_trip = round_trip.commit().unwrap();
        assert!(round_trip.patch().is_empty());
        let round_trip_durable = round_trip.patch().to_durable(patch_limits()).unwrap();
        let digest = exact_digest(source.bytes());
        assert_eq!(
            artifact_preconditions(round_trip_durable.operations()),
            [digest.as_str(), digest.as_str()]
        );
        assert_eq!(
            source.apply_durable(&round_trip_durable).unwrap().bytes(),
            source.bytes()
        );
    }

    #[test]
    fn mixed_formatting_and_structural_patch_binds_the_unretained_intermediate() {
        let source = Snapshot::from_bytes(fixture("45543.ppt")).unwrap();
        let hidden = source.slide_hidden(Position::new(0)).unwrap();
        // The formatting-only publication is exactly the intermediate working
        // artifact that the mixed transaction's structural phase starts from.
        let mut format = source.edit().unwrap();
        format.set_slide_hidden(Position::new(0), !hidden).unwrap();
        let intermediate = format.commit().unwrap().snapshot().clone();
        let intermediate_digest = exact_digest(intermediate.bytes());

        let mut edit = source.edit().unwrap();
        edit.set_slide_hidden(Position::new(0), !hidden).unwrap();
        edit.remove_slide(Position::new(1)).unwrap();
        let commit = edit.commit().unwrap();
        let patch = commit.patch();
        assert_eq!(
            patch.structural_before,
            StructuralArtifact::Intermediate(intermediate_digest.clone())
        );
        assert_eq!(patch.structural_after, StructuralArtifact::After);
        // The retained artifacts are still untouched until serialization.
        assert_eq!(memoized_digest(&patch.before), None);
        assert_eq!(memoized_digest(&patch.after), None);

        let durable = patch.to_durable(transfer_patch_limits()).unwrap();
        let target_digest = exact_digest(commit.snapshot().bytes());
        assert_eq!(
            artifact_preconditions(durable.operations()),
            [intermediate_digest.as_str()]
        );
        assert_eq!(
            artifact_preconditions(durable.inverse().operations()),
            [target_digest.as_str()]
        );
        let inverse = patch.inverse();
        assert_eq!(inverse.structural_before, StructuralArtifact::Before);
        assert_eq!(
            inverse.structural_after,
            StructuralArtifact::Intermediate(intermediate_digest.clone())
        );
        // Removal payloads are normalized by inversion, so compare the
        // artifact bindings rather than the whole double inverse.
        let double = inverse.inverse();
        assert_eq!(double.structural_before, patch.structural_before);
        assert_eq!(double.structural_after, patch.structural_after);
        let inverse_durable = inverse.to_durable(transfer_patch_limits()).unwrap();
        assert_eq!(
            artifact_preconditions(inverse_durable.operations()),
            [target_digest.as_str()]
        );

        let applied = source.apply_durable(&durable).unwrap();
        assert_eq!(applied.bytes(), commit.snapshot().bytes());
        assert_eq!(applied.slide_hidden(Position::new(0)).unwrap(), !hidden);

        // Siblings whose formatting state matches but whose bytes differ:
        // slide 0 is toggled and restored in two commits. The durable
        // formatting operations therefore apply, and only the structural
        // artifact precondition can refuse them.
        let visibility_sibling = |base: &Snapshot, state: bool| {
            let mut away = base.edit().unwrap();
            away.set_slide_hidden(Position::new(0), !state).unwrap();
            let away = away.commit().unwrap().snapshot().clone();
            let mut back = away.edit().unwrap();
            back.set_slide_hidden(Position::new(0), state).unwrap();
            let sibling = back.commit().unwrap().snapshot().clone();
            assert_eq!(sibling.slide_hidden(Position::new(0)).unwrap(), state);
            assert_eq!(sibling.slide_count(), base.slide_count());
            assert_ne!(sibling.bytes(), base.bytes());
            sibling
        };
        // Forward: the hide applies to the sibling, whose intermediate differs
        // from the one the structural removal is bound to.
        let source_sibling = visibility_sibling(&source, hidden);
        assert_structural_artifact_conflict(source_sibling.apply_durable(&durable));
        // Inverse: structural reinsertion runs first against the committed
        // artifact's sibling.
        let target_sibling = visibility_sibling(commit.snapshot(), !hidden);
        assert_structural_artifact_conflict(target_sibling.apply_durable(&durable.inverse()));
        assert_structural_artifact_conflict(target_sibling.apply_durable(&inverse_durable));

        // Complete durable round trip through an unretained intermediate on a
        // deck whose re-inserted slide needs no second picture store.
        let source = Snapshot::from_bytes(authored_fixture()).unwrap();
        let target = crate::text_edit::Target::new(Position::new(0), Position::new(0));
        let original_anchor = source.shape_anchor(target).unwrap();
        let anchor = crate::Anchor::small(25, 35, 325, 235).unwrap();
        let mut edit = source.edit().unwrap();
        edit.set_shape_anchor(target, anchor).unwrap();
        edit.remove_slide(Position::new(1)).unwrap();
        let commit = edit.commit().unwrap();
        assert!(matches!(
            commit.patch().structural_before,
            StructuralArtifact::Intermediate(_)
        ));
        let durable = commit.patch().to_durable(transfer_patch_limits()).unwrap();
        let applied = source.apply_durable(&durable).unwrap();
        assert_eq!(applied.bytes(), commit.snapshot().bytes());
        let restored = applied.apply_durable(&durable.inverse()).unwrap();
        assert_eq!(restored.slide_count(), source.slide_count());
        assert_eq!(slide_texts(restored.bytes()), slide_texts(source.bytes()));
        assert_eq!(restored.shape_anchor(target).unwrap(), original_anchor);

        // The same proof through an anchor that is moved and restored.
        let anchor_sibling = |base: &Snapshot, kept: crate::Anchor, away: crate::Anchor| {
            let mut first = base.edit().unwrap();
            first.set_shape_anchor(target, away).unwrap();
            let moved = first.commit().unwrap().snapshot().clone();
            let mut second = moved.edit().unwrap();
            second.set_shape_anchor(target, kept).unwrap();
            let sibling = second.commit().unwrap().snapshot().clone();
            assert_eq!(sibling.shape_anchor(target).unwrap(), kept);
            assert_ne!(sibling.bytes(), base.bytes());
            sibling
        };
        let source_sibling = anchor_sibling(&source, original_anchor, anchor);
        assert_structural_artifact_conflict(source_sibling.apply_durable(&durable));
        let target_sibling = anchor_sibling(commit.snapshot(), anchor, original_anchor);
        assert_structural_artifact_conflict(target_sibling.apply_durable(&durable.inverse()));
    }

    #[test]
    fn artifact_digest_memo_follows_every_snapshot_construction() {
        let source = Snapshot::from_bytes(authored_batch_fixture()).unwrap();
        let clone = source.clone();
        assert!(source.artifact.shares_memo_with(&clone.artifact));
        assert_eq!(clone.artifact_digest(), exact_digest(source.bytes()));
        assert_eq!(
            memoized_digest(&source),
            Some(exact_digest(source.bytes()).as_str())
        );

        let first = crate::text_edit::Target::new(Position::new(0), Position::new(0));
        let second = crate::text_edit::Target::new(Position::new(0), Position::new(1));
        let mut edit = source.edit().unwrap();
        edit.set_shape_text(first, "single replacement").unwrap();
        let text = edit.working.clone();
        edit.set_shape_texts(&[ShapeTextReplacement::new(second, "batch replacement")])
            .unwrap();
        let batch = edit.working.clone();
        edit.set_shape_anchor(first, crate::Anchor::small(20, 30, 220, 130).unwrap())
            .unwrap();
        let anchor = edit.working.clone();
        for (label, snapshot) in [
            ("single text", &text),
            ("batch text", &batch),
            ("anchor", &anchor),
        ] {
            assert!(
                !source.artifact.shares_memo_with(&snapshot.artifact),
                "{label} publication must not inherit the source memo"
            );
            assert_eq!(memoized_digest(snapshot), None, "{label}");
            assert_eq!(
                snapshot.artifact_digest(),
                exact_digest(snapshot.bytes()),
                "{label}"
            );
        }
        let commit = edit.commit().unwrap();
        assert_eq!(commit.snapshot().bytes(), anchor.bytes());
        assert_eq!(
            commit.snapshot().artifact_digest(),
            exact_digest(anchor.bytes())
        );
    }

    #[test]
    fn snapshot_and_patch_equality_ignore_digest_memo_state() {
        let bytes = authored_fixture();
        let hashed = Snapshot::from_bytes(bytes.clone()).unwrap();
        let fresh = Snapshot::from_bytes(bytes).unwrap();
        let _digest = hashed.artifact_digest();
        assert_eq!(memoized_digest(&fresh), None);
        assert_eq!(hashed, fresh);
        assert_eq!(format!("{hashed:?}"), format!("{fresh:?}"));

        let mut left = hashed.edit().unwrap();
        left.move_slide(Position::new(0), Position::new(1)).unwrap();
        let left = left.commit().unwrap();
        let mut right = fresh.edit().unwrap();
        right
            .move_slide(Position::new(0), Position::new(1))
            .unwrap();
        let right = right.commit().unwrap();
        let _serialized = left.patch().to_durable(patch_limits()).unwrap();
        assert_eq!(left, right);
        assert_eq!(left.patch().inverse().inverse(), *right.patch());
    }

    #[test]
    fn transfer_target_check_reuses_the_memoized_receiver_digest() {
        let donor = Snapshot::from_bytes(authored_fixture()).unwrap();
        let mut remove = donor.edit().unwrap();
        remove.remove_slide(Position::new(1)).unwrap();
        let receiver = remove.commit().unwrap().snapshot().clone();
        assert_eq!(memoized_digest(&receiver), None);
        let plan = receiver
            .plan_transfer_from(&donor, Position::new(0))
            .unwrap();
        assert_eq!(plan.target_artifact, exact_digest(receiver.bytes()));
        assert_eq!(
            memoized_digest(&receiver),
            Some(plan.target_artifact.as_str())
        );
        let mut edit = receiver.edit().unwrap();
        assert!(edit.source.artifact.shares_memo_with(&receiver.artifact));
        edit.insert_transfer(Position::new(1), &plan).unwrap();
        assert_eq!(edit.commit().unwrap().snapshot().slide_count(), 2);

        // A plan for another exact target is still refused before staging.
        let mut stale = donor.edit().unwrap();
        assert!(matches!(
            stale.insert_transfer(Position::new(0), &plan),
            Err(Error::Refused(Refusal::TransferTargetMismatch))
        ));
        assert_eq!(stale.slide_count(), donor.slide_count());
    }

    #[cfg(feature = "performance-diagnostics")]
    fn started_phases(events: &[DiagnosticEvent]) -> Vec<DiagnosticPhase> {
        events
            .iter()
            .filter_map(|event| match event {
                DiagnosticEvent::Started { phase } => Some(*phase),
                DiagnosticEvent::Finished { .. } => None,
            })
            .collect()
    }

    #[cfg(feature = "performance-diagnostics")]
    fn assert_balanced_events(events: &[DiagnosticEvent]) {
        assert_eq!(events.len() % 2, 0);
        for pair in events.chunks_exact(2) {
            let DiagnosticEvent::Started { phase } = pair[0] else {
                panic!("diagnostic phase must start with Started")
            };
            let DiagnosticEvent::Finished {
                phase: finished,
                outcome: _,
            } = pair[1]
            else {
                panic!("diagnostic phase must end with Finished")
            };
            assert_eq!(phase, finished);
        }
    }

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn profiled_commit_reports_the_fixed_phase_order_for_a_real_edit() {
        let source = Snapshot::from_bytes(fixture("45543.ppt")).unwrap();
        let mut ordinary_edit = source.edit().unwrap();
        ordinary_edit
            .remove_slide(Position::new(1))
            .expect("fixture must have a second slide");
        let ordinary = ordinary_edit.commit().unwrap();

        let mut profiled_edit = source.edit().unwrap();
        profiled_edit
            .remove_slide(Position::new(1))
            .expect("fixture must have a second slide");
        let mut events = Vec::new();
        let profiled = profiled_edit
            .commit_profiled(|event| events.push(event))
            .unwrap();

        assert_balanced_events(&events);
        assert_eq!(
            started_phases(&events),
            [
                DiagnosticPhase::DocumentCommit,
                DiagnosticPhase::EmbeddedOpen,
                DiagnosticPhase::BeforePayloadCapture,
                DiagnosticPhase::LiveDocumentRead,
                DiagnosticPhase::EmbeddedFinish,
                DiagnosticPhase::UnrelatedStreamValidation,
                DiagnosticPhase::PublicReopen,
                DiagnosticPhase::AfterPayloadCapture,
            ]
        );
        assert!(events.iter().all(|event| {
            matches!(
                event,
                DiagnosticEvent::Started { .. }
                    | DiagnosticEvent::Finished {
                        outcome: DiagnosticOutcome::Success,
                        ..
                    }
            )
        }));
        assert_eq!(profiled, ordinary);
        assert_eq!(
            profiled.patch().apply(&source).unwrap(),
            *profiled.snapshot()
        );
        assert_eq!(
            profiled
                .patch()
                .inverse()
                .apply(profiled.snapshot())
                .unwrap(),
            source
        );
    }

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn profiled_commit_preserves_formatting_only_output_and_patch() {
        let source = Snapshot::from_bytes(fixture("45543.ppt")).unwrap();
        let original = source.slide_hidden(Position::new(0)).unwrap();

        let mut ordinary_edit = source.edit().unwrap();
        ordinary_edit
            .set_slide_hidden(Position::new(0), !original)
            .unwrap();
        let ordinary = ordinary_edit.commit().unwrap();

        let mut profiled_edit = source.edit().unwrap();
        profiled_edit
            .set_slide_hidden(Position::new(0), !original)
            .unwrap();
        let mut events = Vec::new();
        let profiled = profiled_edit
            .commit_profiled(|event| events.push(event))
            .unwrap();

        assert_balanced_events(&events);
        assert_eq!(profiled, ordinary);
        assert_eq!(
            profiled.patch().apply(&source).unwrap(),
            *profiled.snapshot()
        );
        assert_eq!(
            profiled
                .patch()
                .inverse()
                .apply(profiled.snapshot())
                .unwrap(),
            source
        );
        assert_eq!(
            started_phases(&events),
            [
                DiagnosticPhase::DocumentCommit,
                DiagnosticPhase::StructuralNoOp,
            ]
        );
        assert_ne!(profiled.snapshot().bytes(), source.bytes());
    }

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn profiled_structural_noop_keeps_source_allocation_without_hashing() {
        let source = Snapshot::from_bytes(fixture("45543.ppt")).unwrap();
        let ordinary = source.edit().unwrap().commit().unwrap();
        let mut events = Vec::new();
        let profiled = source
            .edit()
            .unwrap()
            .commit_profiled(|event| events.push(event))
            .unwrap();

        assert_balanced_events(&events);
        assert_eq!(
            started_phases(&events),
            [
                DiagnosticPhase::DocumentCommit,
                DiagnosticPhase::StructuralNoOp,
            ]
        );
        assert_eq!(profiled, ordinary);
        assert!(std::ptr::eq(
            source.bytes().as_ptr(),
            profiled.snapshot().bytes().as_ptr()
        ));
        assert_eq!(source.bytes().len(), profiled.snapshot().bytes().len());
        assert!(profiled.patch().is_empty());
        assert_eq!(
            profiled.patch().apply(&source).unwrap(),
            *profiled.snapshot()
        );
        assert_eq!(
            profiled
                .patch()
                .inverse()
                .apply(profiled.snapshot())
                .unwrap(),
            source
        );
    }

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn profiled_mixed_commit_reports_only_the_intermediate_artifact_hash() {
        let source = Snapshot::from_bytes(fixture("45543.ppt")).unwrap();
        let hidden = source.slide_hidden(Position::new(0)).unwrap();
        let stage = || {
            let mut edit = source.edit().unwrap();
            edit.set_slide_hidden(Position::new(0), !hidden).unwrap();
            edit.remove_slide(Position::new(1)).unwrap();
            edit
        };
        let ordinary = stage().commit().unwrap();
        let mut events = Vec::new();
        let profiled = stage().commit_profiled(|event| events.push(event)).unwrap();

        assert_balanced_events(&events);
        assert_eq!(
            started_phases(&events),
            [
                DiagnosticPhase::DocumentCommit,
                DiagnosticPhase::EmbeddedOpen,
                DiagnosticPhase::BeforePayloadCapture,
                DiagnosticPhase::LiveDocumentRead,
                DiagnosticPhase::EmbeddedFinish,
                DiagnosticPhase::UnrelatedStreamValidation,
                DiagnosticPhase::PublicReopen,
                DiagnosticPhase::AfterPayloadCapture,
                DiagnosticPhase::IntermediateArtifactHash,
            ]
        );
        assert_eq!(profiled, ordinary);
        assert!(matches!(
            profiled.patch().structural_before,
            StructuralArtifact::Intermediate(_)
        ));
        assert_eq!(
            profiled
                .patch()
                .to_durable(transfer_patch_limits())
                .unwrap()
                .to_deterministic_json()
                .unwrap(),
            ordinary
                .patch()
                .to_durable(transfer_patch_limits())
                .unwrap()
                .to_deterministic_json()
                .unwrap()
        );
    }

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn profiled_failure_closes_live_read_with_the_existing_typed_error() {
        let source = Snapshot::from_bytes(fixture("45543.ppt")).unwrap();
        let mut edit = source.edit().unwrap();
        edit.remove_slide(Position::new(1)).unwrap();
        // Keep the staged document candidate while leaving the package bytes
        // at the source artifact. This exercises the existing optimistic
        // live-document guard.
        edit.working.document = edit.document.snapshot().unwrap();
        let mut events = Vec::new();
        let error = edit
            .commit_profiled(|event| events.push(event))
            .unwrap_err();

        assert_balanced_events(&events);
        assert_eq!(
            started_phases(&events),
            [
                DiagnosticPhase::DocumentCommit,
                DiagnosticPhase::EmbeddedOpen,
                DiagnosticPhase::BeforePayloadCapture,
                DiagnosticPhase::LiveDocumentRead,
            ]
        );
        assert!(matches!(
            events.last(),
            Some(DiagnosticEvent::Finished {
                phase: DiagnosticPhase::LiveDocumentRead,
                outcome: DiagnosticOutcome::Error,
            })
        ));
        assert!(matches!(
            error,
            Error::Package(PackageError::Corrupted(message))
                if message.contains("source changed before publication")
        ));
    }
    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn profiled_missing_record_preserves_ordinary_error_and_closes_phase() {
        let source = Snapshot::from_bytes(fixture("45543.ppt")).unwrap();
        let mut edit = source.edit().unwrap();
        edit.remove_slide(Position::new(1)).unwrap();
        // Inject an absent document record at the actual persisted read boundary.
        // The working package and slide payloads remain valid.
        edit.source.document_persist_id = u32::MAX;
        let ordinary = edit.clone().commit().unwrap_err();
        let mut events = Vec::new();
        let profiled = edit
            .commit_profiled(|event| events.push(event))
            .unwrap_err();
        assert_eq!(ordinary.to_string(), profiled.to_string());
        for error in [ordinary, profiled] {
            assert!(
                matches!(error, Error::Package(PackageError::Corrupted(message))
                if message == format!("unknown persist ID {}", u32::MAX))
            );
        }
        assert_balanced_events(&events);
        assert_eq!(
            started_phases(&events),
            [
                DiagnosticPhase::DocumentCommit,
                DiagnosticPhase::EmbeddedOpen,
                DiagnosticPhase::BeforePayloadCapture,
                DiagnosticPhase::LiveDocumentRead,
            ]
        );
        assert_eq!(
            events.last(),
            Some(&DiagnosticEvent::Finished {
                phase: DiagnosticPhase::LiveDocumentRead,
                outcome: DiagnosticOutcome::Error,
            })
        );
    }
}
