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
    UnsupportedSlideDependency { dependency: &'static str },
    /// The receiving presentation does not contain the referenced master.
    MissingMasterDependency { master_id: u32 },
    /// A transfer plan was prepared for a different exact receiving artifact.
    TransferTargetMismatch,
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
            Self::TransferTargetMismatch => {
                formatter.write_str("PPT slide transfer plan belongs to a different target")
            },
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
            text_changes: Vec::new(),
            anchor_changes: Vec::new(),
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
    /// The current conservative closure supports slides whose master identity
    /// already exists in the receiver and which have no notes, drawing-group,
    /// or external-media ownership. Unsupported dependency edges are refused
    /// explicitly instead of copying dangling native identifiers.
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
        let record = persisted_record(donor, donor_slide.persist_id())?;
        require_portable_slide(self, &record)?;
        let payload = slide_payload(group, record)?;
        require_master(&self.document.edit(), payload.master_id)?;
        Ok(TransferPlan {
            target_artifact: artifact_hash(self.bytes()),
            payload: normalized_payload(&payload),
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
}

impl TransferPlan {
    /// Master identity that must already be present in the receiving deck.
    #[must_use]
    pub const fn required_master_id(&self) -> u32 {
        self.payload.master_id
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
    text_changes: Vec<ShapeTextChange>,
    anchor_changes: Vec<ShapeAnchorChange>,
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
        let document_commit = self.document.commit()?;
        let source = self.source;
        let working = self.working;
        if document_commit.patch().is_empty() {
            let patch = Patch::new(
                source,
                working.clone(),
                self.changes,
                self.structural,
                self.text_changes,
                self.anchor_changes,
                self.formatting,
                true,
                artifact_hash(working.bytes()),
                artifact_hash(working.bytes()),
            );
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
        if live.as_slice() != source.document.bytes() {
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
        let patch = Patch::new(
            source,
            snapshot.clone(),
            self.changes,
            self.structural,
            self.text_changes,
            self.anchor_changes,
            self.formatting,
            true,
            artifact_hash(working.bytes()),
            artifact_hash(snapshot.bytes()),
        );
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
    text_changes: Vec<ShapeTextChange>,
    anchor_changes: Vec<ShapeAnchorChange>,
    formatting: Vec<FormattingChange>,
    formatting_first: bool,
    structural_before_artifact: String,
    structural_after_artifact: String,
}

impl Patch {
    fn new(
        before: Snapshot,
        after: Snapshot,
        changes: Vec<Change>,
        structural: Vec<StructuralChange>,
        text_changes: Vec<ShapeTextChange>,
        anchor_changes: Vec<ShapeAnchorChange>,
        formatting: Vec<FormattingChange>,
        formatting_first: bool,
        structural_before_artifact: String,
        structural_after_artifact: String,
    ) -> Self {
        Self {
            before,
            after,
            changes,
            structural,
            text_changes,
            anchor_changes,
            formatting,
            formatting_first,
            structural_before_artifact,
            structural_after_artifact,
        }
    }

    /// Semantic moves represented by this patch.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
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
        Self::new(
            self.after.clone(),
            self.before.clone(),
            self.changes
                .iter()
                .rev()
                .copied()
                .map(Change::inverse)
                .collect(),
            self.structural
                .iter()
                .rev()
                .map(inverse_structural)
                .collect(),
            self.text_changes
                .iter()
                .rev()
                .map(|change| ShapeTextChange {
                    target: change.target,
                    before: change.after.clone(),
                    after: change.before.clone(),
                })
                .collect(),
            self.anchor_changes
                .iter()
                .rev()
                .map(|change| ShapeAnchorChange {
                    target: change.target,
                    before: change.after,
                    after: change.before,
                })
                .collect(),
            self.formatting
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
                })
                .collect(),
            !self.formatting_first,
            self.structural_after_artifact.clone(),
            self.structural_before_artifact.clone(),
        )
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
        "presentation.shape-text.set" | "presentation.shape-anchor.set"
    )
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

fn require_portable_slide(target: &Snapshot, record: &[u8]) -> Result<()> {
    let (root, consumed) = crate::Record::parse_with_limits(record, 0, RecordLimits::default())?;
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
            dependency: "speaker-notes",
        }));
    }
    if !crate::comments::parse_slide_comments(&root)?.is_empty() {
        return Err(Error::Refused(Refusal::UnsupportedSlideDependency {
            dependency: "comment-author-catalog",
        }));
    }
    if contains_record_type(&root, crate::RecordType::InteractiveInfo)
        || contains_record_type(&root, crate::RecordType::TextInteractiveInfoAtom)
    {
        return Err(Error::Refused(Refusal::UnsupportedSlideDependency {
            dependency: "hyperlink-action",
        }));
    }
    let has_drawing = contains_record_type(&root, crate::RecordType::PPDrawing);
    let has_external = contains_record_type(&root, crate::RecordType::ExternalObjectRefAtom);
    if has_drawing || has_external {
        let live_ids = target
            .document
            .slides()
            .iter()
            .map(|slide| slide.persist_id())
            .collect::<BTreeSet<_>>();
        let editor = crate::embedded::object::Editor::open_records_arc_with_limit(
            target.bytes.clone(),
            target.limits.max_package_bytes,
        )?;
        let has_orphaned_closure = editor.persist_ids().into_iter().any(|persist_id| {
            !live_ids.contains(&persist_id)
                && editor
                    .persisted_record(persist_id)
                    .is_ok_and(|candidate| candidate == record)
        });
        if has_orphaned_closure {
            return Ok(());
        }
    }
    if has_drawing {
        return Err(Error::Refused(Refusal::UnsupportedSlideDependency {
            dependency: "drawing-group",
        }));
    }
    if has_external {
        return Err(Error::Refused(Refusal::UnsupportedSlideDependency {
            dependency: "external-media",
        }));
    }
    Ok(())
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
    }
}

fn patch_effects(patch: &Patch) -> BTreeSet<String> {
    let mut effects = BTreeSet::new();
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
    fn transfer_planning_closes_simple_dependencies_and_refuses_drawings() {
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
                dependency: "drawing-group"
            }))
        ));

        let comments = Snapshot::from_bytes(fixture("WithComments.ppt")).unwrap();
        assert!(matches!(
            comments.plan_transfer_from(&comments, Position::new(0)),
            Err(Error::Refused(Refusal::UnsupportedSlideDependency {
                dependency: "comment-author-catalog"
            }))
        ));
    }
}
