//! Unified source-checked editing for opened legacy PPT presentations.
//!
//! `[MS-PPT]` 2.4.14.3 defines presentation order by the sequence of
//! `SlidePersistAtom` records in `SlideListWithTextContainer`. This owner moves
//! each complete slide group, including the outline-text records following its
//! persist atom, and publishes the changed live `DocumentContainer` through a
//! new append-only PPT user edit. The same root composes checked shape text
//! and complete PowerPoint client-anchor geometry formatting. Slide payload
//! records and unrelated CFB streams remain exact unless selected explicitly.

use std::collections::{BTreeMap, BTreeSet};
use std::fmt;
use std::io::Cursor;
use std::sync::Arc;

pub use litchi_core::Position;
pub use litchi_core::patch::{CompositionLimits, HistoryLimits, SubEditJoinFailure};

use crate::document_structure;
use crate::package::{Error as PackageError, Package, RecordLimits};

/// Explicit bounded undo/redo history for slide-order snapshots.
pub type History = litchi_core::patch::History<Snapshot>;

/// Maximum automatic slide-advance delay permitted by `[MS-PPT]` 2.6.6.
pub const MAX_SLIDE_ADVANCE_MS: u32 = 86_399_000;

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
    /// The CFB is signed, encrypted, or contains storage metadata that the
    /// stream-only incremental publisher cannot preserve faithfully.
    UnsupportedSource,
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
    /// The selected slide does not contain exactly one canonical fixed-width
    /// visual transition owner.
    UnsupportedSlideTransitionVisual { position: Position },
    /// A shape edit selected a slide inserted but not yet published.
    UncommittedSlideDependency,
}

impl fmt::Display for Refusal {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::UnsupportedSource => formatter.write_str(
                "PPT source cannot be republished through the safe slide-order transaction",
            ),
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
            Self::UnsupportedSlideTransitionVisual { position } => write!(
                formatter,
                "PPT slide position {} has no unique canonical visual transition atom",
                position.get()
            ),
            Self::UncommittedSlideDependency => formatter
                .write_str("PPT shape text cannot target an inserted slide before publication"),
        }
    }
}

/// Error returned by the opened-presentation slide-order owner.
#[derive(Debug)]
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

#[derive(Clone, PartialEq, Eq)]
struct Lineage(Arc<[u8]>);

impl fmt::Debug for Lineage {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("Lineage(..)")
    }
}

/// Immutable exact whole-package snapshot used by slide-order edits.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
    document: document_structure::Snapshot,
    document_persist_id: u32,
    limits: RecordLimits,
    has_review_history: bool,
}

impl Snapshot {
    /// Opens an exact PPT artifact under default finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the complete package or its live slide directory
    /// is malformed.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes(bytes.to_vec())
    }

    /// Captures an owned exact PPT artifact under default finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the complete package cannot be reopened.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, RecordLimits::default())
    }

    /// Captures an owned exact PPT artifact under explicit finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the artifact exceeds the limits or its live
    /// presentation structure is malformed.
    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: RecordLimits) -> Result<Self> {
        if bytes.len() > limits.max_package_bytes {
            return Err(PackageError::ResourceLimit(format!(
                "PPT slide-order source size {} exceeds limit {}",
                bytes.len(),
                limits.max_package_bytes
            ))
            .into());
        }
        let mut package = Package::from_reader_with_limits(Cursor::new(bytes.as_slice()), limits)?;
        let presentation = package.presentation()?;
        let directory = presentation.slide_directory();
        let (document_persist_id, document_bytes) =
            crate::embedded::object::Editor::inspect_live_document(&bytes)?;
        if document_persist_id != directory.document_persist_id() {
            return Err(PackageError::Corrupted(
                "PPT slide-order snapshot resolved inconsistent live document identities".into(),
            )
            .into());
        }
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
        drop(presentation);
        drop(package);
        Ok(Self {
            bytes: Arc::from(bytes.into_boxed_slice()),
            document,
            document_persist_id,
            limits,
            has_review_history,
        })
    }

    /// Exact bytes of the complete source or committed artifact.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Number of presentation slides in semantic order.
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.document.slides().len()
    }

    /// Whether one slide is omitted from the presentation sequence.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the position is absent.
    pub fn slide_hidden(&self, position: Position) -> Result<bool> {
        self.require_position(position)?;
        Ok(self.document.slides()[position.get()].flags() & (1 << 2) != 0)
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
    pub const fn limits(&self) -> RecordLimits {
        self.limits
    }

    /// Starts an isolated source-bound order edit.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for signed/encrypted or metadata-bearing CFB
    /// sources and for position-dependent reviewer history.
    pub fn edit(&self) -> Result<Transaction> {
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
        })
    }

    /// Prepares one independently composable move against this exact source.
    ///
    /// Moves whose inclusive affected position ranges do not intersect can be
    /// joined; overlapping moves produce the common deterministic conflict
    /// set rather than using last-writer-wins behavior.
    ///
    /// # Errors
    ///
    /// Returns a typed source/position refusal or a common composition bound
    /// error.
    pub fn prepare_move(
        &self,
        limits: CompositionLimits,
        identifier: impl Into<String>,
        from: Position,
        destination: Position,
    ) -> Result<PreparedMove> {
        self.require_editable()?;
        self.require_position(from)?;
        self.require_position(destination)?;
        let move_ = Move { from, destination };
        let start = from.get().min(destination.get());
        let end = from.get().max(destination.get());
        let writes = (start..=end).map(|position| format!("slide-order/position:{position}"));
        let inner = litchi_core::patch::SubEdit::new(
            self.lineage(),
            limits,
            identifier,
            std::iter::empty(),
            writes,
            move_,
        )?;
        Ok(PreparedMove { inner })
    }

    /// Starts a bounded composition of independently prepared moves.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when this source cannot be safely republished.
    pub fn compose(&self, limits: CompositionLimits) -> Result<Composition> {
        self.require_editable()?;
        Ok(Composition {
            source: self.clone(),
            joined: litchi_core::patch::JoinedSubEdits::new(self.lineage(), limits),
        })
    }

    /// Starts explicit bounded undo/redo history at this snapshot.
    #[must_use]
    pub fn history(&self, limits: HistoryLimits) -> History {
        History::new(self.clone(), limits)
    }

    /// Plans a dependency-checked transfer from `donor` without mutating
    /// either presentation.
    ///
    /// The bounded closure reuses a byte-identical master plus matching
    /// comment-author, hyperlink, sound, and external-media owners already
    /// present in the receiver. Ordinary shapes and connector rules are
    /// deterministically rewritten to target-owned unused drawing IDs. Notes,
    /// BLIPs, animation/build graphs, active OLE/macro/program actions, and
    /// missing or mismatched owners are refused explicitly.
    ///
    /// # Errors
    ///
    /// Returns a typed dependency refusal, a position refusal, or a package
    /// error for malformed source records.
    pub fn plan_transfer_from(&self, donor: &Self, slide: Position) -> Result<TransferPlan> {
        self.require_editable()?;
        donor.require_position(slide)?;
        let donor_slide = donor.document.slides()[slide.get()];
        let group = donor.document.edit().slide_group(slide.get())?;
        let mut record = persisted_record(donor, donor_slide.persist_id())?;
        let closure = require_portable_slide(self, donor, &mut record)?;
        let payload = slide_payload(group, record)?;
        require_master_closure(self, donor, payload.master_id)?;
        Ok(TransferPlan {
            target_artifact: artifact_hash(self.bytes()),
            payload: normalized_payload(&payload),
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
    pub fn plan_three_way(&self, left: &Patch, right: &Patch) -> Result<ThreeWayPlan> {
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
    pub fn apply_durable<Mode>(&self, patch: &litchi_core::patch::Patch<Mode>) -> Result<Self> {
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
                    "presentation.shape-anchor.set" => {
                        apply_durable_anchor(&current, &operations[index])?
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
                    _ => unreachable!("formatting vocabulary is checked by its predicate"),
                };
                index += 1;
                continue;
            }

            let structural_artifact = artifact_hash(current.bytes());
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
        if self.has_review_history {
            return Err(Error::Refused(Refusal::ReviewHistoryDependency));
        }
        crate::font::require_stream_only_cfb(self.bytes())
            .and_then(|()| {
                crate::embedded::object::Editor::open_records_arc_with_limit(
                    self.bytes.clone(),
                    self.limits.max_package_bytes,
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

    fn lineage(&self) -> Lineage {
        Lineage(self.bytes.clone())
    }
}

/// One semantic move staged in a transaction or durable patch.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Change {
    from: Position,
    destination: Position,
    before_order: [u8; 32],
    after_order: [u8; 32],
}

impl Change {
    /// Source position resolved against the transaction-local order.
    #[must_use]
    pub const fn from(&self) -> Position {
        self.from
    }

    /// Final destination position for the selected slide.
    #[must_use]
    pub const fn destination(&self) -> Position {
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
pub struct ShapeTextChange {
    target: crate::text_edit::Target,
    before: String,
    after: String,
}

impl ShapeTextChange {
    /// Semantic source-order shape target.
    #[must_use]
    pub const fn target(&self) -> crate::text_edit::Target {
        self.target
    }

    /// Text required before the replacement.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// Replacement text.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// One checked shape-geometry formatting replacement.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ShapeAnchorChange {
    target: crate::text_edit::Target,
    before: crate::Anchor,
    after: crate::Anchor,
}

/// One checked replacement of an inert external movie/audio path.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalMediaPathChange {
    id: u32,
    before: Option<String>,
    after: Option<String>,
}

impl ExternalMediaPathChange {
    /// External-object identifier retained by the replacement.
    #[must_use]
    pub const fn id(&self) -> u32 {
        self.id
    }

    /// Path required before the replacement.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    /// Replacement path, or `None` to clear it.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }
}

/// One checked replacement of external-media playback flags.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExternalMediaPlaybackChange {
    id: u32,
    before: crate::external_media::Playback,
    after: crate::external_media::Playback,
}

/// One fixed-width replacement of a slide's hidden flag.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideVisibilityChange {
    position: Position,
    slide_id: u32,
    before: bool,
    after: bool,
    before_order: [u8; 32],
    after_order: [u8; 32],
}

impl SlideVisibilityChange {
    /// Semantic slide position resolved against the transaction-local order.
    #[must_use]
    pub const fn position(self) -> Position {
        self.position
    }

    /// Hidden state required before the replacement.
    #[must_use]
    pub const fn before(self) -> bool {
        self.before
    }

    /// Replacement hidden state.
    #[must_use]
    pub const fn after(self) -> bool {
        self.after
    }
}

/// One fixed-width replacement of slide-show advance timing and flags.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideAdvanceChange {
    target: Position,
    slide_id: u32,
    before: SlideAdvance,
    after: SlideAdvance,
}

impl SlideAdvanceChange {
    /// Semantic source-order slide target.
    #[must_use]
    pub const fn target(self) -> Position {
        self.target
    }

    /// Advance state required before replacement.
    #[must_use]
    pub const fn before(self) -> SlideAdvance {
        self.before
    }

    /// Replacement advance state.
    #[must_use]
    pub const fn after(self) -> SlideAdvance {
        self.after
    }
}

/// One fixed-width replacement of a slide's visual transition fields.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideTransitionVisualChange {
    target: Position,
    slide_id: u32,
    before: SlideTransitionVisual,
    after: SlideTransitionVisual,
}

impl SlideTransitionVisualChange {
    /// Semantic source-order slide target.
    #[must_use]
    pub const fn target(self) -> Position {
        self.target
    }

    /// Visual transition required before replacement.
    #[must_use]
    pub const fn before(self) -> SlideTransitionVisual {
        self.before
    }

    /// Replacement visual transition.
    #[must_use]
    pub const fn after(self) -> SlideTransitionVisual {
        self.after
    }
}

impl ExternalMediaPlaybackChange {
    /// External-object identifier retained by the replacement.
    #[must_use]
    pub const fn id(self) -> u32 {
        self.id
    }

    /// Playback flags required before the replacement.
    #[must_use]
    pub const fn before(self) -> crate::external_media::Playback {
        self.before
    }

    /// Replacement playback flags.
    #[must_use]
    pub const fn after(self) -> crate::external_media::Playback {
        self.after
    }
}

impl ShapeAnchorChange {
    /// Semantic source-order shape target.
    #[must_use]
    pub const fn target(self) -> crate::text_edit::Target {
        self.target
    }

    /// Anchor required before the replacement.
    #[must_use]
    pub const fn before(self) -> crate::Anchor {
        self.before
    }

    /// Replacement anchor.
    #[must_use]
    pub const fn after(self) -> crate::Anchor {
        self.after
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum FormattingChange {
    Text(ShapeTextChange),
    Anchor(ShapeAnchorChange),
    SlideAdvance(SlideAdvanceChange),
    SlideTransitionVisual(SlideTransitionVisualChange),
    MediaPath(ExternalMediaPathChange),
    MediaPlayback(ExternalMediaPlaybackChange),
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct SlidePayload {
    group: Vec<crate::records::Record>,
    record: Vec<u8>,
    master_id: u32,
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
pub struct ThreeWayConflict {
    target: String,
}

impl ThreeWayConflict {
    /// Stable semantic target shared by both branches.
    #[must_use]
    pub fn target(&self) -> &str {
        &self.target
    }
}

/// Read-only conflict result for two patches prepared from one exact base.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ThreeWayPlan {
    conflicts: Vec<ThreeWayConflict>,
}

impl ThreeWayPlan {
    fn new(left: &Patch, right: &Patch) -> Self {
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
    pub fn is_clean(&self) -> bool {
        self.conflicts.is_empty()
    }

    /// Stable sorted conflict set.
    #[must_use]
    pub fn conflicts(&self) -> &[ThreeWayConflict] {
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
    Visibility(SlideVisibilityChange),
    Remove(ListChange),
    Insert(ListChange),
}

/// Isolated failure-atomic edit over one opened presentation's slide order.
#[derive(Debug, Clone)]
pub struct Transaction {
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
}

impl Transaction {
    /// Immutable source snapshot.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Current number of slides.
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.document
            .slides()
            .map_or(self.source.slide_count(), |slides| slides.len())
    }

    /// Semantic moves staged in call order.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Fixed-width slide visibility changes staged in call order.
    #[must_use]
    pub fn slide_visibility_changes(&self) -> &[SlideVisibilityChange] {
        &self.visibility_changes
    }

    /// Fixed-width slide advance changes staged in call order.
    #[must_use]
    pub fn slide_advance_changes(&self) -> &[SlideAdvanceChange] {
        &self.advance_changes
    }

    /// Fixed-width visual transition changes staged in call order.
    #[must_use]
    pub fn slide_transition_visual_changes(&self) -> &[SlideTransitionVisualChange] {
        &self.transition_visual_changes
    }

    /// Checked shape-text replacements staged in this root transaction.
    #[must_use]
    pub fn shape_text_changes(&self) -> &[ShapeTextChange] {
        &self.text_changes
    }

    /// Checked shape-anchor formatting replacements staged in this root.
    #[must_use]
    pub fn shape_anchor_changes(&self) -> &[ShapeAnchorChange] {
        &self.anchor_changes
    }

    /// External movie/audio path replacements staged in this root.
    #[must_use]
    pub fn external_media_path_changes(&self) -> &[ExternalMediaPathChange] {
        &self.media_path_changes
    }

    /// External-media playback replacements staged in this root.
    #[must_use]
    pub fn external_media_playback_changes(&self) -> &[ExternalMediaPlaybackChange] {
        &self.media_playback_changes
    }

    /// Moves a slide to a final zero-based position in the current projected
    /// order. Moving a slide to its current position is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when either position is absent, or a package
    /// error when the structural candidate is invalid.
    pub fn move_slide(&mut self, from: Position, destination: Position) -> Result<()> {
        self.require_position(from)?;
        self.require_position(destination)?;
        if from == destination {
            return Ok(());
        }
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

    /// Changes whether one slide is hidden during presentation.
    ///
    /// Only bit 2 of the live `SlidePersistAtom.flags` field is changed. The
    /// operation composes with order, shape, and inert media changes and is
    /// carried by exact, inverse, durable, history, and merge surfaces.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the position is absent or a package error
    /// when the fixed-width document mutation fails validation.
    pub fn set_slide_hidden(&mut self, position: Position, hidden: bool) -> Result<()> {
        self.require_position(position)?;
        let selected = self.document.slides()?[position.get()];
        let before = selected.flags() & (1 << 2) != 0;
        if before == hidden {
            return Ok(());
        }
        let before_order = order_digest(&self.document)?;
        self.document.set_slide_hidden(position.get(), hidden)?;
        let after_order = order_digest(&self.document)?;
        let change = SlideVisibilityChange {
            position,
            slide_id: selected.slide_id(),
            before,
            after: hidden,
            before_order,
            after_order,
        };
        self.visibility_changes.push(change);
        self.structural.push(StructuralChange::Visibility(change));
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
    pub fn set_slide_advance(&mut self, position: Position, advance: SlideAdvance) -> Result<()> {
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
    pub fn set_slide_transition_visual(
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
    pub fn set_shape_text(
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
        let text_snapshot = crate::text_edit::Snapshot::from_bytes(self.working.bytes().to_vec())?;
        let mut text_edit = text_snapshot.edit_text(source_target)?;
        let before = text_edit.text().to_string();
        let after = value.into();
        text_edit.set_text(after.clone())?;
        let commit = text_edit.commit()?;
        if commit.patch().is_empty() {
            return Ok(());
        }
        self.working = Snapshot::from_bytes_with_limits(
            commit.snapshot().bytes().to_vec(),
            self.source.limits,
        )?;
        let change = ShapeTextChange {
            target: source_target,
            before,
            after,
        };
        self.text_changes.push(change.clone());
        self.formatting.push(FormattingChange::Text(change));
        Ok(())
    }

    /// Replaces an existing shape's complete checked `PowerPoint` client anchor.
    /// Every other `OfficeArt` record and CFB stream remains exact.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the semantic target is missing,
    /// ambiguous, newly inserted, or lacks exactly one supported host anchor.
    pub fn set_shape_anchor(
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
        let rewritten =
            crate::text_edit::replace_shape_anchor(self.working.bytes(), source_target, anchor)?;
        if rewritten.before == rewritten.after {
            return Ok(());
        }
        self.working = Snapshot::from_bytes_with_limits(rewritten.bytes, self.source.limits)?;
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
    pub fn set_external_media_path(&mut self, id: u32, path: Option<String>) -> Result<()> {
        let media = external_media_snapshot(&self.working)?;
        let object = media
            .collection()
            .and_then(|collection| collection.objects.iter().find(|object| object.id() == id))
            .ok_or_else(|| {
                PackageError::InvalidFormat(format!("external media ID {id} was not found"))
            })?;
        let before = media_path_of(object)?;
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
    pub fn set_external_media_playback(
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
    pub fn remove_slide(&mut self, position: Position) -> Result<()> {
        self.require_position(position)?;
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

    /// Inserts a dependency-checked cross-presentation transfer plan.
    /// `position == slide_count()` appends the transferred slide.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for a stale/different target or a package error
    /// when identifier allocation or structural validation fails.
    pub fn insert_transfer(&mut self, position: Position, plan: &TransferPlan) -> Result<()> {
        if plan.target_artifact != litchi_core::patch::BlobId::of(self.source.bytes()).as_hex() {
            return Err(Error::Refused(Refusal::TransferTargetMismatch));
        }
        if position.get() > self.slide_count() {
            return Err(Error::Refused(Refusal::SlideNotFound { position }));
        }
        require_master(&self.document, plan.payload.master_id)?;
        let persist_id = next_persist_id(&self.source, &self.inserted_records)?;
        let slide_id = next_slide_id(&self.document)?;
        let mut group = plan.payload.group.clone();
        rewrite_slide_persist(&mut group, persist_id, slide_id)?;
        let before_order = order_digest(&self.document)?;
        self.document
            .insert_slide_group(position.get(), group.clone())?;
        let after_order = order_digest(&self.document)?;
        self.inserted_records
            .insert(persist_id, plan.payload.record.clone());
        self.structural.push(StructuralChange::Insert(ListChange {
            position,
            before_order,
            after_order,
            payload: SlidePayload {
                group,
                record: plan.payload.record.clone(),
                master_id: plan.payload.master_id,
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
    pub fn commit(self) -> Result<Commit> {
        let document_commit = if self.structural.is_empty()
            || (self.media_path_changes.is_empty() && self.media_playback_changes.is_empty())
        {
            self.document.commit()?
        } else {
            let mut rebased = self.working.document.edit();
            replay_structural_changes(&mut rebased, &self.structural)?;
            rebased.commit()?
        };
        let source = self.source;
        let working = self.working;
        if document_commit.patch().is_empty() {
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
                structural_before_artifact: artifact_hash(working.bytes()),
                structural_after_artifact: artifact_hash(working.bytes()),
            };
            return Ok(Commit {
                snapshot: working,
                patch,
            });
        }

        let before_slides = persisted_slides(&working.bytes, &working.document)?;
        let mut editor = crate::embedded::object::Editor::open_records_arc_with_limit(
            working.bytes.clone(),
            source.limits.max_package_bytes,
        )?;
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
        editor.replace_persisted_record(
            source.document_persist_id,
            document_commit.snapshot().bytes().to_vec(),
        )?;
        let bytes = editor.finish()?;
        crate::font::validate_unrelated_streams(working.bytes(), &bytes)?;
        let snapshot = Snapshot::from_bytes_with_limits(bytes, source.limits)?;
        if snapshot.document.slides() != document_commit.snapshot().slides() {
            return Err(PackageError::Corrupted(
                "published PPT slide order did not round-trip through the live document".into(),
            )
            .into());
        }
        let after_slides = persisted_slides(&snapshot.bytes, &snapshot.document)?;
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
            structural_before_artifact: artifact_hash(working.bytes()),
            structural_after_artifact: artifact_hash(snapshot.bytes()),
        };
        Ok(Commit { snapshot, patch })
    }

    /// Discards the candidate and recovers the exact source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    fn require_position(&self, position: Position) -> Result<()> {
        if position.get() >= self.slide_count() {
            return Err(Error::Refused(Refusal::SlideNotFound { position }));
        }
        Ok(())
    }

    fn order_fingerprint(&self) -> Result<String> {
        order_digest(&self.document).map(hex_digest)
    }
}

/// Successful immutable target and its reversible source-checked patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Published whole-package snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible patch from the exact source to the committed artifact.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Splits the commit into its target and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Reversible exact-source-checked whole-package slide-order patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
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
    structural_before_artifact: String,
    structural_after_artifact: String,
}

impl Patch {
    /// Semantic moves represented by this patch.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Slide hidden-state replacements composed into this root patch.
    #[must_use]
    pub fn slide_visibility_changes(&self) -> &[SlideVisibilityChange] {
        &self.visibility_changes
    }

    /// Slide advance replacements composed into this root patch.
    #[must_use]
    pub fn slide_advance_changes(&self) -> &[SlideAdvanceChange] {
        &self.advance_changes
    }

    /// Visual transition replacements composed into this root patch.
    #[must_use]
    pub fn slide_transition_visual_changes(&self) -> &[SlideTransitionVisualChange] {
        &self.transition_visual_changes
    }

    /// Shape-text replacements composed into this root patch.
    #[must_use]
    pub fn shape_text_changes(&self) -> &[ShapeTextChange] {
        &self.text_changes
    }

    /// Shape-anchor formatting replacements composed into this root patch.
    #[must_use]
    pub fn shape_anchor_changes(&self) -> &[ShapeAnchorChange] {
        &self.anchor_changes
    }

    /// External movie/audio path replacements composed into this root patch.
    #[must_use]
    pub fn external_media_path_changes(&self) -> &[ExternalMediaPathChange] {
        &self.media_path_changes
    }

    /// External-media playback replacements composed into this root patch.
    #[must_use]
    pub fn external_media_playback_changes(&self) -> &[ExternalMediaPlaybackChange] {
        &self.media_playback_changes
    }

    /// Whether this patch changes no artifact bytes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.bytes() == self.after.bytes()
    }

    /// Exact source artifact required for forward application.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        self.before.bytes()
    }

    /// Exact committed target artifact.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        self.after.bytes()
    }

    /// Applies only to the exact source artifact used to build this patch.
    ///
    /// # Errors
    ///
    /// Returns an error when `current` is not the exact source snapshot.
    pub fn apply(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.bytes() != self.before.bytes() {
            return Err(PackageError::InvalidFormat(
                "PPT slide-order patch source does not match its base artifact".into(),
            )
            .into());
        }
        Ok(self.after.clone())
    }

    /// Returns the exact-source-checked inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
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
                    slide_id: change.slide_id,
                    before: change.after,
                    after: change.before,
                    before_order: change.after_order,
                    after_order: change.before_order,
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
                    FormattingChange::Anchor(anchor_change) => {
                        FormattingChange::Anchor(ShapeAnchorChange {
                            target: anchor_change.target,
                            before: anchor_change.after,
                            after: anchor_change.before,
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
                })
                .collect(),
            formatting_first: !self.formatting_first,
            structural_before_artifact: self.structural_after_artifact.clone(),
            structural_after_artifact: self.structural_before_artifact.clone(),
        }
    }

    /// Converts this patch to the common bounded deterministic-JSON envelope.
    ///
    /// Operations use only semantic zero-based positions, exact artifact
    /// SHA-256, and a content-free order fingerprint; native slide and persist
    /// identifiers are never serialized.
    ///
    /// # Errors
    ///
    /// Returns an error when caller-selected patch bounds cannot represent the
    /// semantic operations.
    pub fn to_durable(
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
                        &self.structural_before_artifact,
                        move_change.before_order,
                    )?,
                    durable_operation(
                        limits,
                        move_change.destination,
                        move_change.from,
                        &self.structural_after_artifact,
                        move_change.after_order,
                    )?,
                ),
                StructuralChange::Visibility(visibility_change) => (
                    durable_visibility_operation(
                        limits,
                        visibility_change,
                        &self.structural_before_artifact,
                        visibility_change.before_order,
                    )?,
                    durable_visibility_operation(
                        limits,
                        &SlideVisibilityChange {
                            position: visibility_change.position,
                            slide_id: visibility_change.slide_id,
                            before: visibility_change.after,
                            after: visibility_change.before,
                            before_order: visibility_change.after_order,
                            after_order: visibility_change.before_order,
                        },
                        &self.structural_after_artifact,
                        visibility_change.after_order,
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
                            &self.structural_before_artifact,
                            list_change.before_order,
                        )?,
                        durable_insert_operation(
                            limits,
                            list_change,
                            &self.structural_after_artifact,
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
                            &self.structural_before_artifact,
                            list_change.before_order,
                            &forward_blob.as_hex(),
                        )?,
                        durable_remove_operation(
                            limits,
                            list_change,
                            &self.structural_after_artifact,
                            list_change.after_order,
                        )?,
                    )
                },
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Move {
    from: Position,
    destination: Position,
}

/// One independently prepared move retaining its exact source lineage.
pub struct PreparedMove {
    inner: litchi_core::patch::SubEdit<Lineage, Move>,
}

impl PreparedMove {
    /// Stable caller-selected composition identifier.
    #[must_use]
    pub fn identifier(&self) -> &str {
        self.inner.id()
    }

    /// Source position in the immutable base order.
    #[must_use]
    pub const fn from(&self) -> Position {
        self.inner.payload().from
    }

    /// Final destination in the immutable base order.
    #[must_use]
    pub const fn destination(&self) -> Position {
        self.inner.payload().destination
    }
}

impl fmt::Debug for PreparedMove {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PreparedMove")
            .field("identifier", &self.identifier())
            .field("from", &self.from())
            .field("destination", &self.destination())
            .finish()
    }
}

/// Recoverable failure to join one prepared slide move.
pub struct JoinError {
    failure: Box<SubEditJoinFailure>,
    rejected: Box<PreparedMove>,
}

impl JoinError {
    /// Structured common composition refusal.
    #[must_use]
    pub const fn failure(&self) -> &SubEditJoinFailure {
        &self.failure
    }

    /// Recovers the rejected prepared move.
    #[must_use]
    pub fn into_rejected(self) -> PreparedMove {
        *self.rejected
    }
}

impl fmt::Debug for JoinError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("JoinError")
            .field("failure", &self.failure)
            .field("rejected", &self.rejected)
            .finish()
    }
}

/// Bounded deterministic composition of provably disjoint slide moves.
pub struct Composition {
    source: Snapshot,
    joined: litchi_core::patch::JoinedSubEdits<Lineage, Move>,
}

impl Composition {
    /// Joins a move when its source lineage and affected position range are
    /// compatible with every accepted move.
    ///
    /// # Errors
    ///
    /// Returns the structured common refusal and retains the rejected work.
    pub fn join(&mut self, incoming: PreparedMove) -> std::result::Result<&mut Self, JoinError> {
        if let Err(error) = self.joined.join(incoming.inner) {
            let (failure, rejected) = error.into_parts();
            return Err(JoinError {
                failure: Box::new(failure),
                rejected: Box::new(PreparedMove { inner: rejected }),
            });
        }
        Ok(self)
    }

    /// Number of accepted independently prepared moves.
    #[must_use]
    pub fn len(&self) -> usize {
        self.joined.len()
    }

    /// Whether no move has been accepted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.joined.is_empty()
    }

    /// Atomically commits every accepted move in deterministic identifier
    /// order and returns one durable-capable patch.
    ///
    /// # Errors
    ///
    /// Returns an error if staging or whole-package publication fails.
    pub fn commit(self) -> Result<Commit> {
        let mut transaction = self.source.edit()?;
        for edit in self.joined.into_sub_edits() {
            let move_ = edit.into_payload();
            transaction.move_slide(move_.from, move_.destination)?;
        }
        transaction.commit()
    }
}

impl fmt::Debug for Composition {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Composition")
            .field("moves", &self.joined.len())
            .finish_non_exhaustive()
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
    change: &SlideVisibilityChange,
    artifact: &str,
    order: [u8; 32],
) -> std::result::Result<litchi_core::patch::PatchOperation, litchi_core::patch::PatchError> {
    let mut preconditions = structural_preconditions(artifact, order);
    preconditions.insert("hidden".to_string(), serde_json::Value::Bool(change.before));
    litchi_core::patch::PatchOperation::new(
        limits,
        "presentation.slide-hidden.set",
        format!("slide:{}", change.position.get()),
        preconditions,
        serde_json::Value::Bool(change.after),
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
    let text_snapshot = crate::text_edit::Snapshot::from_bytes(snapshot.bytes().to_vec())?;
    let mut edit = text_snapshot.edit_text(target)?;
    if artifact_hash(edit.text().as_bytes()) != expected {
        return Err(PackageError::InvalidFormat(
            "PPT durable shape-text semantic precondition does not match".into(),
        )
        .into());
    }
    edit.set_text(replacement)?;
    let committed = edit.commit()?;
    Snapshot::from_bytes_with_limits(committed.snapshot().bytes().to_vec(), snapshot.limits)
}

fn is_formatting_operation(operation: &str) -> bool {
    matches!(
        operation,
        "presentation.shape-text.set"
            | "presentation.shape-anchor.set"
            | "presentation.slide-advance.set"
            | "presentation.slide-transition-visual.set"
            | "presentation.external-media-path.set"
            | "presentation.external-media-playback.set"
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
    if media_path_digest(snapshot.external_media_path(id)?.as_deref()) != expected {
        return Err(PackageError::InvalidFormat(
            "PPT durable external-media path precondition does not match".into(),
        )
        .into());
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
    if playback_digest(snapshot.external_media_playback(id)?) != expected {
        return Err(PackageError::InvalidFormat(
            "PPT durable external-media playback precondition does not match".into(),
        )
        .into());
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
    let current = crate::text_edit::inspect_shape_anchor(snapshot.bytes(), target)?;
    if artifact_hash(&current.to_bytes()) != expected {
        return Err(PackageError::InvalidFormat(
            "PPT durable shape-anchor semantic precondition does not match".into(),
        )
        .into());
    }
    let replacement = parse_anchor_value(&operation.value)?;
    let rewritten = crate::text_edit::replace_shape_anchor(snapshot.bytes(), target, replacement)?;
    Snapshot::from_bytes_with_limits(rewritten.bytes, snapshot.limits)
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
    replace_slide_advance(
        snapshot,
        target,
        parse_slide_advance_value(&operation.value)?,
    )
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
    replace_slide_transition_visual(
        snapshot,
        target,
        parse_transition_visual_value(&operation.value)?,
    )
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
        "presentation.slide-hidden.set" if operation.preconditions.len() == 3 => {
            let expected_hidden = operation
                .preconditions
                .get("hidden")
                .and_then(serde_json::Value::as_bool)
                .ok_or_else(|| invalid_durable_patch("missing slide-hidden precondition"))?;
            transaction.require_position(position)?;
            let current_hidden =
                transaction.document.slides()?[position.get()].flags() & (1 << 2) != 0;
            if current_hidden != expected_hidden {
                return Err(PackageError::InvalidFormat(
                    "PPT durable slide-hidden semantic precondition does not match".into(),
                )
                .into());
            }
            let replacement = operation
                .value
                .as_bool()
                .ok_or_else(|| invalid_durable_patch("slide-hidden value must be boolean"))?;
            transaction.set_slide_hidden(position, replacement)
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
    for (index, pair) in hex.as_bytes().chunks_exact(2).enumerate() {
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

fn external_media_snapshot(snapshot: &Snapshot) -> Result<crate::external_media::Snapshot> {
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
        snapshot.bytes.clone(),
        snapshot.limits.max_package_bytes,
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
    let published = Snapshot::from_bytes_with_limits(bytes, snapshot.limits)?;
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
) -> Result<BTreeMap<u32, Vec<u8>>> {
    let editor = crate::embedded::object::Editor::open_records_arc_with_limit(
        bytes.clone(),
        bytes.len().saturating_mul(4),
    )?;
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
    crate::embedded::object::Editor::open_records_arc_with_limit(
        snapshot.bytes.clone(),
        snapshot.limits.max_package_bytes,
    )?
    .persisted_record(persist_id)
    .map_err(Error::from)
}

fn inspect_slide_advance(snapshot: &Snapshot, position: Position) -> Result<SlideAdvance> {
    let slide = snapshot
        .document
        .slides()
        .get(position.get())
        .copied()
        .ok_or(Error::Refused(Refusal::SlideNotFound { position }))?;
    let bytes = persisted_record(snapshot, slide.persist_id())?;
    let (root, consumed) = crate::Record::parse_with_limits(&bytes, 0, snapshot.limits)?;
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
    let (mut root, consumed) =
        crate::Record::parse_with_limits(&source_record, 0, snapshot.limits)?;
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
        snapshot.bytes.clone(),
        snapshot.limits.max_package_bytes,
    )?;
    editor.replace_persisted_record(slide.persist_id(), rewritten_record)?;
    let bytes = editor.finish()?;
    crate::font::validate_unrelated_streams(snapshot.bytes(), &bytes)?;
    let published = Snapshot::from_bytes_with_limits(bytes, snapshot.limits)?;
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
    let (root, consumed) = crate::Record::parse_with_limits(&bytes, 0, snapshot.limits)?;
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
    let (mut root, consumed) =
        crate::Record::parse_with_limits(&source_record, 0, snapshot.limits)?;
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
        snapshot.bytes.clone(),
        snapshot.limits.max_package_bytes,
    )?;
    editor.replace_persisted_record(slide.persist_id(), rewritten_record)?;
    let bytes = editor.finish()?;
    crate::font::validate_unrelated_streams(snapshot.bytes(), &bytes)?;
    let published = Snapshot::from_bytes_with_limits(bytes, snapshot.limits)?;
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
        source.bytes.clone(),
        source.limits.max_package_bytes,
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
    let (root, consumed) = crate::Record::parse_with_limits(&record, 0, RecordLimits::default())?;
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
    })
}

struct TransferClosure {
    dependencies: Vec<TransferDependency>,
    remaps: Vec<RelationshipRemap>,
}

fn require_portable_slide(
    target: &Snapshot,
    donor: &Snapshot,
    record: &mut Vec<u8>,
) -> Result<TransferClosure> {
    let (mut root, consumed) =
        crate::Record::parse_with_limits(record, 0, RecordLimits::default())?;
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
    let comments = crate::comments::parse_slide_comments(&root)?;
    if !comments.is_empty() {
        let target_authors = crate::comments::Authors::parse(target.document.record())?;
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
    })
}

fn relationship_remaps(
    hyperlinks: &BTreeMap<u32, u32>,
    sounds: &BTreeMap<u32, u32>,
    external_objects: &BTreeMap<u32, u32>,
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
    shape_ids: BTreeSet<u32>,
    connector_shape_ids: BTreeSet<u32>,
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

fn option_has_blip_reference(instance: u16, payload: &[u8]) -> Result<bool> {
    let property_bytes = usize::from(instance)
        .checked_mul(6)
        .ok_or_else(|| PackageError::Corrupted("OfficeArt property count overflow".into()))?;
    let properties = payload
        .get(..property_bytes)
        .ok_or_else(|| PackageError::Corrupted("OfficeArt property table is truncated".into()))?;
    Ok(properties
        .chunks_exact(6)
        .any(|property| u16::from_le_bytes([property[0], property[1]]) & 0x4000 != 0))
}

fn option_has_linked_shape_reference(instance: u16, payload: &[u8]) -> Result<bool> {
    let property_bytes = usize::from(instance)
        .checked_mul(6)
        .ok_or_else(|| PackageError::Corrupted("OfficeArt property count overflow".into()))?;
    let properties = payload
        .get(..property_bytes)
        .ok_or_else(|| PackageError::Corrupted("OfficeArt property table is truncated".into()))?;
    Ok(properties.chunks_exact(6).any(|property| {
        let opid = u16::from_le_bytes([property[0], property[1]]) & 0x3fff;
        opid == 0x008a && property[2..6].iter().any(|byte| *byte != 0)
    }))
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

fn drawing_remaps(target: &Snapshot, donor: &SlideRelationships) -> Result<BTreeMap<u32, u32>> {
    if donor
        .unsupported_dependencies
        .contains(&TransferDependency::PictureStore)
    {
        return Err(Error::Refused(Refusal::UnsupportedSlideDependency {
            dependency: TransferDependency::PictureStore,
        }));
    }
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
        target.bytes.clone(),
        target.limits.max_package_bytes,
    )?;
    let mut live_shape_ids = BTreeSet::new();
    let mut available_shape_ids = BTreeSet::new();
    for persist_id in editor.persist_ids() {
        let Ok(bytes) = editor.persisted_record(persist_id) else {
            continue;
        };
        let Ok((root, consumed)) = crate::Record::parse_with_limits(&bytes, 0, target.limits)
        else {
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
                    if option_has_blip_reference(instance, payload)? {
                        output
                            .unsupported_dependencies
                            .insert(TransferDependency::PictureStore);
                    }
                    if option_has_linked_shape_reference(instance, payload)? {
                        output
                            .unsupported_dependencies
                            .insert(TransferDependency::OtherOfficeArtShapeReference);
                    }
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
                0xF014 | 0xF017 | 0xF11D | 0xF119 => {
                    output
                        .unsupported_dependencies
                        .insert(TransferDependency::OtherOfficeArtShapeReference);
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
    let donor_links = crate::Hyperlinks::parse(donor.document.record())?;
    let target_links = crate::Hyperlinks::parse(target.document.record())?;
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
            drawings,
        )?;
    }
    for child in &mut root.children {
        rewrite_slide_relationships(child, hyperlinks, sounds, external_objects, drawings)?;
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
    drawings: &BTreeMap<u32, u32>,
) -> Result<()> {
    visit_raw_records_mut(
        bytes,
        "OfficeArt drawing",
        |version, instance, kind, payload| {
            match kind {
                0xF00A => rewrite_u32_reference(payload, drawings, "OfficeArtSp.spid")?,
                0xF012 if version == 1 && instance == 0 => {
                    rewrite_connector_rule(payload, drawings)?;
                },
                0xF012 => {
                    return Err(PackageError::Corrupted(
                        "OfficeArtFConnectorRule has an invalid record header".into(),
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
                        drawings,
                    )?;
                },
                _ => {},
            }
            Ok(())
        },
    )
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
        StructuralChange::Visibility(change) => {
            StructuralChange::Visibility(SlideVisibilityChange {
                position: change.position,
                slide_id: change.slide_id,
                before: change.after,
                after: change.before,
                before_order: change.after_order,
                after_order: change.before_order,
            })
        },
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
            StructuralChange::Visibility(change) => {
                let current = document.slides()?[change.position.get()].flags() & (1 << 2) != 0;
                if current != change.before {
                    return Err(PackageError::Corrupted(
                        "PPT structural rebase selected a different slide visibility state".into(),
                    )
                    .into());
                }
                document.set_slide_hidden(change.position.get(), change.after)?;
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
        }
    }
    Ok(())
}

fn patch_effects(patch: &Patch) -> BTreeSet<String> {
    let mut effects = BTreeSet::new();
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
            StructuralChange::Visibility(change) => {
                effects.insert(format!("slide-id:{}/hidden", change.slide_id));
            },
            StructuralChange::Remove(list_change) | StructuralChange::Insert(list_change) => {
                effects.insert(format!(
                    "slide-list/position:{}",
                    list_change.position.get()
                ));
                if let Ok((root, consumed)) = crate::Record::parse_with_limits(
                    &list_change.payload.record,
                    0,
                    RecordLimits::default(),
                ) && consumed == list_change.payload.record.len()
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
        }
    }
    effects
}

fn encode_payload(payload: &SlidePayload) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"LSP1");
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
    Ok(bytes)
}

fn decode_payload(bytes: &[u8], limits: RecordLimits) -> Result<SlidePayload> {
    if bytes.len() > limits.max_input_bytes || bytes.get(0..4) != Some(b"LSP1") {
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
        let (record, consumed) = crate::Record::parse_with_limits(encoded, 0, limits)?;
        if consumed != encoded.len() {
            return Err(invalid_durable_patch(
                "slide group blob record has trailing bytes",
            ));
        }
        group.push(record);
    }
    let record = take_length_prefixed(bytes, &mut offset, limits.max_record_bytes)?.to_vec();
    if offset != bytes.len() {
        return Err(invalid_durable_patch(
            "slide insertion blob has trailing bytes",
        ));
    }
    let payload = slide_payload(group, record)?;
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

    fn fixture(name: &str) -> Vec<u8> {
        std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/poi/test-data/slideshow")
                .join(name),
        )
        .unwrap()
    }

    fn slide_texts(bytes: &[u8]) -> Vec<String> {
        let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
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
        let mut writer = crate::writer::Writer::new();
        let first = writer.add_slide().unwrap();
        writer
            .add_textbox(first, 10, 10, 240, 40, "first slide")
            .unwrap();
        let second = writer.add_slide().unwrap();
        writer
            .add_textbox(second, 10, 10, 240, 40, "second slide")
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    fn authored_table_fixture() -> Vec<u8> {
        let mut writer = crate::writer::Writer::new();
        let slide = writer.add_slide().unwrap();
        let mut table = crate::writer::Table::new(2, 2).unwrap();
        table.set_cell_text(0, 0, "A1").unwrap();
        table.set_cell_text(1, 1, "B2").unwrap();
        writer.add_table(slide, 50, 60, table).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
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
        let (root, consumed) =
            crate::Record::parse_with_limits(&bytes, 0, snapshot.limits).unwrap();
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
    fn real_fixture_reorders_reopens_and_round_trips_durable_history() {
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

        let mut history = source.history(HistoryLimits::new(4, 4096));
        history
            .record(commit.snapshot().clone(), wire.len() as u64)
            .unwrap();
        assert!(history.undo());
        assert_eq!(history.current(), &source);
        assert!(history.redo());
        assert_eq!(history.current(), commit.snapshot());
    }

    #[test]
    fn common_composition_joins_disjoint_moves_and_refuses_overlap() {
        let source = Snapshot::from_bytes(fixture("45543.ppt")).unwrap();
        let limits = CompositionLimits::new(8, 32, 64, 32);
        let first = source
            .prepare_move(
                limits,
                "move-opening-pair",
                Position::new(0),
                Position::new(1),
            )
            .unwrap();
        let second = source
            .prepare_move(
                limits,
                "move-later-pair",
                Position::new(4),
                Position::new(5),
            )
            .unwrap();
        let overlapping = source
            .prepare_move(
                limits,
                "overlap-opening-pair",
                Position::new(1),
                Position::new(2),
            )
            .unwrap();
        let mut composition = source.compose(limits).unwrap();
        composition.join(first).unwrap().join(second).unwrap();
        let failure = composition.join(overlapping).unwrap_err();
        assert!(matches!(failure.failure(), SubEditJoinFailure::Overlap(_)));
        let commit = composition.commit().unwrap();
        assert_eq!(commit.snapshot().slide_count(), source.slide_count());
        assert_eq!(commit.patch().changes().len(), 2);
        let durable = commit.patch().to_durable(patch_limits()).unwrap();
        assert_eq!(source.apply_durable(&durable).unwrap(), *commit.snapshot());
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
        let mut reopened = Package::from_reader(Cursor::new(commit.snapshot().bytes())).unwrap();
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
        let mut reopened = Package::from_reader(Cursor::new(commit.snapshot().bytes())).unwrap();
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
            Package::from_reader(Cursor::new(media_first_commit.snapshot().bytes())).unwrap();
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
        let mut reopened = Package::from_reader(Cursor::new(commit.snapshot().bytes())).unwrap();
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
    fn connector_picture_store_and_animation_build_refusals_are_distinct() {
        let source = Snapshot::from_bytes(fixture("basic_test_ppt_file.ppt")).unwrap();
        let mut linked_shape_property = Vec::new();
        linked_shape_property.extend_from_slice(&0x008a_u16.to_le_bytes());
        linked_shape_property.extend_from_slice(&0x0401_u32.to_le_bytes());
        assert!(option_has_linked_shape_reference(1, &linked_shape_property).unwrap());
        linked_shape_property[0..2].copy_from_slice(&0x4001_u16.to_le_bytes());
        assert!(option_has_blip_reference(1, &linked_shape_property).unwrap());

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

        let picture = SlideRelationships {
            unsupported_dependencies: BTreeSet::from([TransferDependency::PictureStore]),
            ..SlideRelationships::default()
        };
        assert!(matches!(
            drawing_remaps(&source, &picture),
            Err(Error::Refused(Refusal::UnsupportedSlideDependency {
                dependency: TransferDependency::PictureStore
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

    #[test]
    fn slide_visibility_composes_with_order_and_round_trips_durably() {
        let source = Snapshot::from_bytes(fixture("basic_test_ppt_file.ppt")).unwrap();
        let original = source.slide_hidden(Position::new(0)).unwrap();
        let mut edit = source.edit().unwrap();
        edit.set_slide_hidden(Position::new(0), !original).unwrap();
        edit.move_slide(Position::new(0), Position::new(1)).unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(commit.patch().slide_visibility_changes().len(), 1);
        assert_eq!(
            commit.snapshot().slide_hidden(Position::new(1)).unwrap(),
            !original
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
        let mut reopened = Package::from_reader(Cursor::new(visibility_restored.bytes())).unwrap();
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
    fn genuine_corpus_slide_advance_reopens_reverses_merges_and_uses_history() {
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
            let mut reopened =
                Package::from_reader(Cursor::new(commit.snapshot().bytes())).unwrap();
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

        let mut history = source.history(HistoryLimits::new(4, 64 * 1024));
        history
            .record(commit.snapshot().clone(), wire.len() as u64)
            .unwrap();
        assert!(history.undo());
        assert_eq!(history.current(), &source);
        assert!(history.redo());
        assert_eq!(history.current(), commit.snapshot());

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
}
