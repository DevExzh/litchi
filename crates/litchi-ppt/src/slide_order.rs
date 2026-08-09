//! Source-checked slide ordering for opened legacy PPT presentations.
//!
//! `[MS-PPT]` 2.4.14.3 defines presentation order by the sequence of
//! `SlidePersistAtom` records in `SlideListWithTextContainer`. This owner moves
//! each complete slide group, including the outline-text records following its
//! persist atom, and publishes the changed live `DocumentContainer` through a
//! new append-only PPT user edit. Slide payload records and unrelated CFB
//! streams remain exact.

use std::collections::BTreeMap;
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
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package(error) => error.fmt(formatter),
            Self::Refused(refusal) => refusal.fmt(formatter),
            Self::Composition(error) => error.fmt(formatter),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Package(error) => Some(error),
            Self::Composition(error) => Some(error),
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
            document: self.document.edit(),
            changes: Vec::new(),
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

    /// Applies supported durable slide-order operations to this exact source.
    ///
    /// # Errors
    ///
    /// Returns an error for an unsupported vocabulary, malformed selector,
    /// failed exact-artifact/order precondition, or refused move.
    pub fn apply_durable<Mode>(&self, patch: &litchi_core::patch::Patch<Mode>) -> Result<Self> {
        if patch.format() != "litchi-ppt" || !patch.blobs().is_empty() {
            return Err(invalid_durable_patch("unsupported format or blob bundle"));
        }
        if patch.operations().is_empty() {
            return Ok(self.clone());
        }
        let expected_artifact = litchi_core::patch::BlobId::of(self.bytes()).as_hex();
        for operation in patch.operations() {
            if operation.op != "slide-order.move" || operation.preconditions.len() != 2 {
                return Err(invalid_durable_patch(
                    "unsupported slide-order operation vocabulary",
                ));
            }
            let artifact = operation
                .preconditions
                .get("artifact_sha256")
                .and_then(serde_json::Value::as_str)
                .ok_or_else(|| invalid_durable_patch("missing artifact hash precondition"))?;
            if artifact != expected_artifact {
                return Err(PackageError::InvalidFormat(
                    "PPT durable slide-order patch source artifact does not match".into(),
                )
                .into());
            }
        }

        let mut transaction = self.edit()?;
        for operation in patch.operations() {
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
            let from = parse_target(&operation.target)?;
            let destination = operation
                .value
                .as_u64()
                .and_then(|value| usize::try_from(value).ok())
                .map(Position::new)
                .ok_or_else(|| invalid_durable_patch("slide destination must fit usize"))?;
            transaction.move_slide(from, destination)?;
        }
        transaction.commit().map(|commit| commit.snapshot)
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

/// Isolated failure-atomic edit over one opened presentation's slide order.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    document: document_structure::Transaction,
    changes: Vec<Change>,
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
        self.source.slide_count()
    }

    /// Semantic moves staged in call order.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
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
        self.changes.push(Change {
            from,
            destination,
            before_order,
            after_order,
        });
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
        if document_commit.patch().is_empty() {
            let patch = Patch::new(source.clone(), source.clone(), Vec::new());
            return Ok(Commit {
                snapshot: source,
                patch,
            });
        }

        let before_slides = persisted_slides(&source.bytes, &source.document)?;
        let mut editor = crate::embedded::object::Editor::open_records_arc_with_limit(
            source.bytes.clone(),
            source.limits.max_package_bytes,
        )?;
        let live = editor.persisted_record(source.document_persist_id)?;
        if live.as_slice() != source.document.bytes() {
            return Err(PackageError::Corrupted(
                "PPT slide-order transaction source changed before publication".into(),
            )
            .into());
        }
        editor.replace_persisted_record(
            source.document_persist_id,
            document_commit.snapshot().bytes().to_vec(),
        )?;
        let bytes = editor.finish()?;
        crate::font::validate_unrelated_streams(source.bytes(), &bytes)?;
        let snapshot = Snapshot::from_bytes_with_limits(bytes, source.limits)?;
        if snapshot.document.slides() != document_commit.snapshot().slides() {
            return Err(PackageError::Corrupted(
                "published PPT slide order did not round-trip through the live document".into(),
            )
            .into());
        }
        if persisted_slides(&snapshot.bytes, &snapshot.document)? != before_slides {
            return Err(PackageError::Corrupted(
                "PPT slide-order publication changed a slide payload record".into(),
            )
            .into());
        }
        let patch = Patch::new(source, snapshot.clone(), self.changes);
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
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot, changes: Vec<Change>) -> Self {
        Self {
            before,
            after,
            changes,
        }
    }

    /// Semantic moves represented by this patch.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
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

        let forward_artifact = litchi_core::patch::BlobId::of(self.before()).as_hex();
        let reverse_artifact = litchi_core::patch::BlobId::of(self.after()).as_hex();
        let operations = self
            .changes
            .iter()
            .map(|change| {
                let forward = durable_operation(
                    limits,
                    change.from,
                    change.destination,
                    &forward_artifact,
                    change.before_order,
                )?;
                let inverse = durable_operation(
                    limits,
                    change.destination,
                    change.from,
                    &reverse_artifact,
                    change.after_order,
                )?;
                Ok(ReversibleOperation::new(forward, inverse))
            })
            .collect::<std::result::Result<Vec<_>, litchi_core::patch::PatchError>>()?;
        litchi_core::patch::Patch::<litchi_core::patch::Reversible>::new(
            limits,
            "litchi-ppt",
            operations,
            BlobBundle::new(limits.blobs()),
            BlobBundle::new(limits.blobs()),
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
}
