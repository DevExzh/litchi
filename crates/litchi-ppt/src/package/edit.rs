//! The canonical primary transaction for an immutable PPT [`Package`].
//!
//! The package edit is an adapter around the established, lossless
//! [`crate::slide_order`] owner.  The native transaction remains private to
//! this module: callers get one stable `Edit`/`Patch`/`Commit` vocabulary and
//! an immutable package is returned by every publication.  Each staged call
//! first runs against a clone of the native transaction.  That extra clone is
//! deliberate.  A few native operations perform validation after touching
//! their working document; publishing the candidate only after the complete
//! call succeeds keeps the canonical transaction failure-atomic.

use super::{Error as PackageError, OpenOptions, Package};
pub use crate::document_structure::CustomTableStylesPlacement;
pub use crate::external_media::Playback;
use crate::slide_order;
pub use crate::slide_order::{
    ShapeTextReplacement, SlideAdvance, SlideTransitionVisual, TransferPlan,
};
pub use crate::text_edit::Target;
pub use litchi_core::Position;
use litchi_core::SourceVersion;
use std::fmt;

use crate::scoped::Lazy;

/// The finite operation dimension of the primary PPT edit.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Number of successful staged edit operations.
    Operations,
}

/// The default and hard bound for one primary PPT edit.
///
/// 16384 since record 0264 (owner approved 2026-10-14; it was 4096).
pub const DEFAULT_MAX_OPERATIONS: usize = 16384;
/// The maximum operation budget accepted by [`EditOptions`].
pub const MAX_OPERATIONS: usize = DEFAULT_MAX_OPERATIONS;

/// Checked bounds for a primary PPT edit.
///
/// The fields remain private so an edit can never be constructed with an
/// unchecked budget.  The default is also the hard bound: callers may choose
/// a stricter budget, but cannot widen the finite safety policy.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct EditOptions {
    max_operations: usize,
}

impl fmt::Debug for EditOptions {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("EditOptions")
            .field("max_operations", &self.max_operations)
            .finish()
    }
}

impl Default for EditOptions {
    fn default() -> Self {
        Self::new()
    }
}

impl EditOptions {
    /// Construct the checked default profile.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_operations: DEFAULT_MAX_OPERATIONS,
        }
    }

    /// Select a stricter finite operation budget.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] for zero or a value above the hard
    /// bound.
    pub fn max_operations(mut self, value: usize) -> Result<Self> {
        if value == 0 || value > MAX_OPERATIONS {
            return Err(Error::InvalidLimit {
                kind: LimitKind::Operations,
                value,
                maximum: MAX_OPERATIONS,
            });
        }
        self.max_operations = value;
        Ok(self)
    }

    /// Return the selected operation budget.
    #[must_use]
    pub const fn max_operations_value(&self) -> usize {
        self.max_operations
    }
}

/// Typed errors returned by the canonical PPT primary transaction.
#[non_exhaustive]
pub enum Error {
    /// The immutable package rejected an ingress, publication, or freshness
    /// check.
    Package(PackageError),
    /// The established native PPT owner rejected a semantic operation.
    Native(slide_order::Error),
    /// A requested edit budget was outside its checked range.
    InvalidLimit {
        /// Dimension whose value was invalid.
        kind: LimitKind,
        /// Supplied value.
        value: usize,
        /// Hard maximum permitted for this dimension.
        maximum: usize,
    },
    /// A successful call would exceed the selected edit budget.
    Limit {
        /// Dimension whose bound was exceeded.
        kind: LimitKind,
        /// Number that would be retained after the call.
        observed: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// A patch's exact source image does not match the receiving package.
    PatchConflict,
    /// A changed write would invalidate the package signature.
    Signed,
    /// A durable patch could not be built within the selected bounds
    /// (record 0357a).
    Durable(litchi_core::patch::PatchError),
}

impl From<PackageError> for Error {
    fn from(error: PackageError) -> Self {
        Self::Package(error)
    }
}

impl From<slide_order::Error> for Error {
    fn from(error: slide_order::Error) -> Self {
        match error {
            slide_order::Error::Refused(slide_order::Refusal::Signed) => Self::Signed,
            // A transfer plan made against another receiving artifact is a
            // patch for another base (record 0351 §2.4).
            slide_order::Error::Refused(slide_order::Refusal::TransferTargetMismatch) => {
                Self::PatchConflict
            },
            error => Self::Native(error),
        }
    }
}

impl fmt::Debug for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        // Leaf errors contain parser-owned strings.  Keep this public
        // wrapper content-free while retaining the typed leaf in `source()`.
        match self {
            Self::Package(_) => formatter.write_str("Error::Package(..)"),
            Self::Native(_) => formatter.write_str("Error::Native(..)"),
            Self::InvalidLimit {
                kind,
                value,
                maximum,
            } => formatter
                .debug_struct("Error::InvalidLimit")
                .field("kind", kind)
                .field("value", value)
                .field("maximum", maximum)
                .finish(),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => formatter
                .debug_struct("Error::Limit")
                .field("kind", kind)
                .field("observed", observed)
                .field("maximum", maximum)
                .finish(),
            Self::PatchConflict => formatter.write_str("Error::PatchConflict"),
            Self::Signed => formatter.write_str("Error::Signed"),
            Self::Durable(_) => formatter.write_str("Error::Durable(..)"),
        }
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Signed => {
                formatter.write_str("changed PPT writes require explicit signature removal")
            },
            Self::Package(_) => formatter.write_str("PPT package operation failed"),
            Self::Native(_) => formatter.write_str("PPT native edit operation failed"),
            Self::InvalidLimit {
                kind,
                value,
                maximum,
            } => write!(
                formatter,
                "invalid PPT edit limit for {kind:?}: {value} (maximum {maximum})"
            ),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "PPT edit limit for {kind:?}: {observed} exceeds {maximum}"
            ),
            Self::PatchConflict => formatter.write_str(
                "PPT package patch source does not match the exact package receiving it",
            ),
            Self::Durable(_) => formatter.write_str("PPT durable patch conversion failed"),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Package(error) => Some(error),
            Self::Native(error) => Some(error),
            Self::Durable(error) => Some(error),
            Self::InvalidLimit { .. } | Self::Limit { .. } | Self::PatchConflict | Self::Signed => {
                None
            },
        }
    }
}

impl Error {
    /// Return the limit dimension, if this is a limit failure.
    #[must_use]
    pub const fn limit_kind(&self) -> Option<LimitKind> {
        match self {
            Self::InvalidLimit { kind, .. } | Self::Limit { kind, .. } => Some(*kind),
            Self::Package(_)
            | Self::Native(_)
            | Self::PatchConflict
            | Self::Signed
            | Self::Durable(_) => None,
        }
    }

    /// Return the observed and maximum values, if this is a limit failure.
    #[must_use]
    pub const fn limit_values(&self) -> Option<(usize, usize)> {
        match self {
            Self::InvalidLimit { value, maximum, .. }
            | Self::Limit {
                observed: value,
                maximum,
                ..
            } => Some((*value, *maximum)),
            Self::Package(_)
            | Self::Native(_)
            | Self::PatchConflict
            | Self::Signed
            | Self::Durable(_) => None,
        }
    }
}

/// Result type for the canonical PPT primary transaction.
pub type Result<T> = std::result::Result<T, Error>;

/// Content-free facts about one primary PPT publication.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    operation_count: usize,
}

impl Diagnostics {
    fn new(changed: bool, operation_count: usize) -> Self {
        Self {
            changed,
            operation_count,
        }
    }

    /// Whether exact package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of accepted operation inputs, including exact no-ops.
    /// Each item in a text batch counts separately; failed calls charge nothing.
    #[must_use]
    pub const fn operation_count(self) -> usize {
        self.operation_count
    }
}

/// Exact, reversible primary PPT package patch.
#[derive(Clone)]
pub struct Patch {
    inner: slide_order::Patch,
    operation_count: usize,
    /// Built by a durable replay: exact bytes, no semantic operations.
    replayed: bool,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before_bytes", &self.before().len())
            .field("after_bytes", &self.after().len())
            .field("operation_count", &self.operation_count)
            .field("is_noop", &self.is_noop())
            .finish()
    }
}

impl Patch {
    fn from_native(inner: slide_order::Patch, operation_count: usize) -> Self {
        Self {
            inner,
            operation_count,
            replayed: false,
        }
    }

    /// Whether exact package bytes are identical.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.inner.is_empty()
    }

    /// Exact source artifact required before application.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        self.inner.before()
    }

    /// Exact target artifact produced by application.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        self.inner.after()
    }

    /// Number of accepted operation inputs retained by this patch.
    #[must_use]
    pub const fn operation_count(&self) -> usize {
        self.operation_count
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            inner: self.inner.inverse(),
            operation_count: self.operation_count,
            replayed: self.replayed,
        }
    }

    /// Convert this patch to the shared durable patch envelope
    /// (`litchi_core::patch`), replayed by [`Package::apply_durable`].
    ///
    /// The operations carry semantic zero-based positions, the SHA-256 of
    /// the exact artifacts they bind and a content-free order fingerprint;
    /// native slide and persist identifiers are never serialized. Every
    /// primary-edit verb has a durable form; master reorders, custom
    /// table-style edits and `fNonOutlineData` are the
    /// `document-structure.*` operations (record 0357b), bound like slide
    /// moves, with the custom table-style atoms as blobs.
    ///
    /// # Errors
    ///
    /// [`Error::Durable`] when `limits` cannot hold the operations (a
    /// custom table-style atom needs a blob), and
    /// [`slide_order::Refusal::UnsupportedPatch`] for a changed patch that a
    /// durable replay produced (it keeps the exact bytes, not the
    /// operations; convert the durable patch that was replayed instead).
    pub fn to_durable(
        &self,
        limits: litchi_core::patch::PatchLimits,
    ) -> Result<litchi_core::patch::Patch<litchi_core::patch::Reversible>> {
        if self.replayed && !self.is_noop() {
            return Err(Error::Native(slide_order::Error::Refused(
                slide_order::Refusal::UnsupportedPatch {
                    reason: "a replayed durable patch keeps no semantic operations",
                },
            )));
        }
        self.inner.to_durable(limits).map_err(Error::Durable)
    }

    /// Join an independently prepared patch over the same exact source.
    ///
    /// The two patches must commute (records 0357a, 0357b). Each may hold
    /// slide moves, master reorders, custom table-style edits,
    /// `fNonOutlineData` changes and the formatting edits (shape and notes
    /// text, anchors, slide visibility, advance and visual transition,
    /// external media path and playback). Neither patch's slide-list
    /// positions (moved ranges, `fNonOutlineData` slides) may meet the
    /// other's moved positions, and no semantic target may be written by
    /// both. The joined patch stages both patches' formatting on the
    /// source, then `self`'s structure and `other`'s, in one commit, and
    /// counts the operations of both.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when the patches start from different
    /// packages; [`slide_order::Refusal::UnsupportedPatch`] (kind
    /// `UnsupportedOperation`) for slide insertions or removals, element
    /// record edits, an inverse patch with both formatting and structure,
    /// or positions that meet a moved range; and
    /// [`slide_order::Refusal::JoinConflict`] (kind `UnsupportedOperation`)
    /// for a target both patches write, such as two master reorders.
    pub fn join(&self, other: &Self) -> Result<Self> {
        if self.before() != other.before() {
            return Err(Error::PatchConflict);
        }
        if self.replayed || other.replayed {
            return Err(Error::Native(slide_order::Error::Refused(
                slide_order::Refusal::UnsupportedPatch {
                    reason: "a replayed durable patch keeps no semantic operations",
                },
            )));
        }
        let operation_count = self
            .operation_count
            .checked_add(other.operation_count)
            .ok_or(Error::Limit {
                kind: LimitKind::Operations,
                observed: usize::MAX,
                maximum: MAX_OPERATIONS,
            })?;
        let inner = self.inner.join(&other.inner)?;
        Ok(Self::from_native(inner, operation_count))
    }
}

/// One semantic target written by both patches of a [`ThreeWayPlan`].
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Conflict {
    target: String,
}

impl Conflict {
    /// Stable semantic target written by both patches, such as
    /// `slide-id:256/transition-visual` or `slide-list/position:1`.
    #[must_use]
    pub fn target(&self) -> &str {
        &self.target
    }
}

/// Read-only three-way plan of two patches made from one exact package.
///
/// It lists the semantic targets both patches write; a clean plan has none.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ThreeWayPlan {
    conflicts: Vec<Conflict>,
}

impl ThreeWayPlan {
    /// Whether the two patches write no common semantic target.
    #[must_use]
    pub fn is_clean(&self) -> bool {
        self.conflicts.is_empty()
    }

    /// The common semantic targets, sorted.
    #[must_use]
    pub fn conflicts(&self) -> &[Conflict] {
        &self.conflicts
    }
}

/// Successful immutable PPT primary publication.
#[derive(Clone)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("package", &self.package)
            .field("patch", &self.patch)
            .field("diagnostics", &self.diagnostics)
            .finish()
    }
}

impl Commit {
    /// Published immutable package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Exact reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Content-free publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Consume the commit into its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit into its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

/// Detached primary PPT edit backed by the checked native slide-order owner.
pub struct Edit {
    package: Package,
    inner: Lazy<slide_order::Transaction>,
    options: EditOptions,
    operation_count: usize,
    expected_source: SourceVersion,
}

impl fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("package", &self.package)
            .field("options", &self.options)
            .field("operation_count", &self.operation_count)
            .finish_non_exhaustive()
    }
}

impl Edit {
    pub(crate) fn from_package_with_options(
        package: Package,
        options: EditOptions,
    ) -> Result<Self> {
        if package.is_encrypted() {
            return Err(Error::Package(PackageError::EncryptionPolicy));
        }
        let expected_source = package.source_version().map_err(Error::Package)?;
        Ok(Self {
            package,
            inner: Lazy::new(),
            options,
            operation_count: 0,
            expected_source,
        })
    }

    fn initialize_inner(&self) -> Result<slide_order::Transaction> {
        let bytes = self.package.to_encrypted_bytes().map_err(Error::Package)?;
        ensure_source_version(&self.package, self.expected_source)?;
        let snapshot =
            slide_order::Snapshot::from_bytes_with_options(bytes, open_options(&self.package));
        ensure_source_version(&self.package, self.expected_source)?;
        let inner = snapshot?.edit();
        ensure_source_version(&self.package, self.expected_source)?;
        // The package's own projection was parsed from exactly these bytes;
        // target resolution reuses its parsed slides (record 0260).
        Ok(inner?.with_opened_projection(self.package.presentation().map_err(Error::Package)?))
    }

    fn inner(&self) -> Result<&slide_order::Transaction> {
        self.inner.get_or_try_init(|| self.initialize_inner())
    }

    fn inner_mut(&mut self) -> Result<&mut slide_order::Transaction> {
        let package = self.package.clone();
        let expected_source = self.expected_source;
        self.inner.get_or_try_init_mut(|| {
            let bytes = package.to_encrypted_bytes().map_err(Error::Package)?;
            ensure_source_version(&package, expected_source)?;
            let snapshot =
                slide_order::Snapshot::from_bytes_with_options(bytes, open_options(&package));
            ensure_source_version(&package, expected_source)?;
            let inner = snapshot?.edit();
            ensure_source_version(&package, expected_source)?;
            Ok(inner?.with_opened_projection(package.presentation().map_err(Error::Package)?))
        })
    }

    /// Immutable package that authorized this edit.
    #[must_use]
    pub const fn source(&self) -> &Package {
        &self.package
    }

    /// Selected checked edit profile.
    #[must_use]
    pub const fn options(&self) -> EditOptions {
        self.options
    }

    /// Number of accepted operation inputs, counting text batch items separately.
    #[must_use]
    pub const fn operation_count(&self) -> usize {
        self.operation_count
    }

    /// Move one slide to a final zero-based position.
    pub fn move_slide(&mut self, from: Position, destination: Position) -> Result<()> {
        self.stage(1, |inner| inner.move_slide(from, destination))
    }

    /// Remove one slide and its order-owned outline records.
    pub fn remove_slide(&mut self, position: Position) -> Result<()> {
        self.stage(1, |inner| inner.remove_slide(position))
    }

    /// Insert a blank slide at a projected position (`slide_count()`
    /// appends). It follows the master and color scheme of its preceding
    /// slide and receives a new drawing and shape-identifier cluster.
    pub fn insert_slide(&mut self, position: Position) -> Result<()> {
        self.stage(1, |inner| inner.insert_slide(position))
    }

    /// Insert a dependency-checked transfer plan at a projected position.
    pub fn insert_transfer(&mut self, position: Position, plan: &TransferPlan) -> Result<()> {
        self.stage(1, |inner| inner.insert_transfer(position, plan))
    }

    /// Set one existing shape's text.
    pub fn set_shape_text(&mut self, target: Target, value: impl Into<String>) -> Result<()> {
        self.stage(1, move |inner| inner.set_shape_text(target, value.into()))
    }

    /// Set several existing shape texts atomically.
    pub fn set_shape_texts(&mut self, replacements: &[ShapeTextReplacement]) -> Result<usize> {
        let charge = replacements.len();
        self.stage(charge, |inner| inner.set_shape_texts(replacements))
    }

    /// Set shape texts from several bounded chunks as one atomic publication.
    ///
    /// Every chunk retains the 256-replacement contract of
    /// [`Self::set_shape_texts`]. The native owner admits at most 4096 chunks,
    /// including empty chunks, and at most 16384 replacements across the
    /// whole call. Targets must be globally distinct. The operation count charged
    /// to this primary edit is the aggregate replacement count, so a valid
    /// chunked call consumes the same finite operation budget as its flattened
    /// equivalent. Empty plans consume no operations.
    ///
    /// # Errors
    ///
    /// Returns the canonical operation-limit error before native work when
    /// the aggregate charge exceeds this edit's selected budget. Native
    /// chunk, replacement-byte, duplicate-target, encoding, dependency,
    /// signed-source, and publication refusals leave this edit unchanged.
    pub fn set_shape_text_chunks(&mut self, chunks: &[&[ShapeTextReplacement]]) -> Result<usize> {
        if chunks.len() > slide_order::MAX_SHAPE_TEXT_CHUNKS {
            ensure_source_version(&self.package, self.expected_source)?;
            return Err(Error::Package(PackageError::ResourceLimit(format!(
                "PPT shape-text chunk count {} exceeds limit {}",
                chunks.len(),
                slide_order::MAX_SHAPE_TEXT_CHUNKS
            ))));
        }
        let charge = chunks.iter().try_fold(0usize, |total, chunk| {
            total.checked_add(chunk.len()).ok_or_else(|| {
                Error::Package(PackageError::ResourceLimit(
                    "PPT shape-text aggregate operation count overflow".into(),
                ))
            })
        })?;
        self.stage(charge, |inner| inner.set_shape_text_chunks(chunks))
    }

    /// Set the speaker-notes text of the slide at a zero-based position.
    ///
    /// The slide's notes page must hold exactly one body placeholder. Equal
    /// text is an exact no-op; a signed source refuses a changed text.
    pub fn set_notes_text(&mut self, position: Position, value: impl Into<String>) -> Result<()> {
        self.stage(1, move |inner| inner.set_notes_text(position, value.into()))
    }

    /// Set one existing shape's complete checked client anchor.
    pub fn set_shape_anchor(&mut self, target: Target, anchor: crate::Anchor) -> Result<()> {
        self.stage(1, |inner| inner.set_shape_anchor(target, anchor))
    }

    /// Set one slide's hidden state.
    pub fn set_slide_hidden(&mut self, position: Position, hidden: bool) -> Result<()> {
        self.stage(1, |inner| inner.set_slide_hidden(position, hidden))
    }

    /// Set one slide's manual/automatic advance state.
    pub fn set_slide_advance(&mut self, position: Position, advance: SlideAdvance) -> Result<()> {
        self.stage(1, |inner| inner.set_slide_advance(position, advance))
    }

    /// Set one slide's visual transition state.
    pub fn set_slide_transition_visual(
        &mut self,
        position: Position,
        visual: SlideTransitionVisual,
    ) -> Result<()> {
        self.stage(1, |inner| {
            inner.set_slide_transition_visual(position, visual)
        })
    }

    /// Set or clear one inert external media path.
    pub fn set_external_media_path(&mut self, id: u32, path: Option<String>) -> Result<()> {
        self.stage(1, |inner| inner.set_external_media_path(id, path))
    }

    /// Set one external media playback policy.
    pub fn set_external_media_playback(&mut self, id: u32, playback: Playback) -> Result<()> {
        self.stage(1, |inner| inner.set_external_media_playback(id, playback))
    }

    /// Append a comment to the slide at a zero-based position; returns its
    /// zero-based position among the slide's comments (record 0323).
    pub fn add_comment(
        &mut self,
        position: Position,
        comment: &crate::comments::NewComment,
    ) -> Result<Position> {
        self.stage(1, |inner| inner.add_comment(position, comment))
            .map(Position::new)
    }

    /// Change the text or anchor of one slide comment (record 0323).
    pub fn set_comment(
        &mut self,
        position: Position,
        comment: Position,
        change: &crate::comments::CommentChange,
    ) -> Result<()> {
        self.stage(1, |inner| inner.set_comment(position, comment, change))
    }

    /// Remove one slide comment (record 0323).
    pub fn remove_comment(&mut self, position: Position, comment: Position) -> Result<()> {
        self.stage(1, |inner| inner.remove_comment(position, comment))
    }

    /// Replace the text of one table cell (record 0323).
    pub fn set_table_cell_text(
        &mut self,
        target: Target,
        row: Position,
        column: Position,
        text: &str,
    ) -> Result<()> {
        self.stage(1, |inner| {
            inner.set_table_cell_text(target, row, column, text)
        })
    }

    /// Set or remove the click hyperlink of one top-level shape (record
    /// 0323).
    pub fn set_shape_hyperlink(
        &mut self,
        target: Target,
        click: Option<&crate::hyperlink::Click>,
    ) -> Result<()> {
        self.stage(1, |inner| inner.set_shape_hyperlink(target, click))
    }

    /// Apply character formatting to a UTF-16 range of one top-level
    /// shape's text (record 0323).
    pub fn set_run_format(
        &mut self,
        target: Target,
        start: usize,
        length: usize,
        format: &crate::text_edit::RunFormat,
    ) -> Result<()> {
        self.stage(1, |inner| {
            inner.set_run_format(target, start, length, format)
        })
    }

    /// Set or remove the click hyperlink over a UTF-16 range of one
    /// top-level shape's text (record 0323).
    pub fn set_text_hyperlink(
        &mut self,
        target: Target,
        start: usize,
        length: usize,
        click: Option<&crate::hyperlink::Click>,
    ) -> Result<()> {
        self.stage(1, |inner| {
            inner.set_text_hyperlink(target, start, length, click)
        })
    }

    /// Move one main or title master to a final zero-based position in the
    /// document's master list (record 0357a).
    ///
    /// Slides, notes and title masters reference a master by its `masterId`,
    /// which is unchanged, so every reference keeps its master.
    ///
    /// # Errors
    ///
    /// [`slide_order::Refusal::ElementNotFound`] for a position outside the
    /// master list; [`Error::Signed`] for a change on a signed package.
    pub fn move_master(&mut self, from: Position, destination: Position) -> Result<()> {
        self.stage(1, |inner| inner.move_master(from, destination))
    }

    /// Move the master whose `masterId` is `master_id` to a final position
    /// (record 0357a).
    ///
    /// # Errors
    ///
    /// [`slide_order::Refusal::ElementNotFound`] when no master has that id.
    pub fn move_master_id(&mut self, master_id: u32, destination: Position) -> Result<()> {
        self.stage(1, |inner| inner.move_master_id(master_id, destination))
    }

    /// Reorder the master list: it becomes `order`, each entry a current
    /// master position (record 0357a). The identity order is a no-op.
    ///
    /// # Errors
    ///
    /// [`slide_order::Refusal::InvalidElementValue`] unless `order` is a
    /// permutation of the master positions.
    pub fn reorder_masters(&mut self, order: &[Position]) -> Result<()> {
        let order: Vec<usize> = order.iter().map(|position| position.get()).collect();
        self.stage(1, |inner| inner.reorder_masters(&order))
    }

    /// Insert a `RoundTripCustomTableStyles12Atom` (its complete record
    /// bytes: the 8-byte header and the payload) before or after the
    /// `EndDocumentAtom` (record 0357a).
    ///
    /// # Errors
    ///
    /// [`slide_order::Refusal::InvalidElementValue`] for bytes that are not
    /// exactly one such atom; [`slide_order::Refusal::UnsupportedElement`]
    /// when the document already has one.
    pub fn insert_custom_table_styles(
        &mut self,
        atom: &[u8],
        placement: CustomTableStylesPlacement,
    ) -> Result<()> {
        let record = custom_table_styles_record(atom)?;
        self.stage(1, |inner| {
            inner.insert_custom_table_styles(record, placement)
        })
    }

    /// Remove the custom table-style atom; returns its complete record bytes
    /// (record 0357a).
    ///
    /// # Errors
    ///
    /// [`slide_order::Refusal::ElementNotFound`] when the document has none.
    pub fn remove_custom_table_styles(&mut self) -> Result<Vec<u8>> {
        let record = self.stage(1, slide_order::Transaction::remove_custom_table_styles)?;
        encode_atom(&record)
    }

    /// Replace the custom table-style atom in place with `atom` (its complete
    /// record bytes); returns the previous atom's bytes (record 0357a). An
    /// equal atom is a no-op.
    ///
    /// # Errors
    ///
    /// [`slide_order::Refusal::ElementNotFound`] when the document has none;
    /// [`slide_order::Refusal::InvalidElementValue`] for bytes that are not
    /// exactly one such atom.
    pub fn replace_custom_table_styles(&mut self, atom: &[u8]) -> Result<Vec<u8>> {
        let record = custom_table_styles_record(atom)?;
        let before = self.stage(1, |inner| inner.replace_custom_table_styles(record))?;
        encode_atom(&before)
    }

    /// Move the custom table-style atom before or after the
    /// `EndDocumentAtom` (record 0357a). Its current placement is a no-op.
    ///
    /// # Errors
    ///
    /// [`slide_order::Refusal::ElementNotFound`] when the document has none.
    pub fn move_custom_table_styles(
        &mut self,
        placement: CustomTableStylesPlacement,
    ) -> Result<()> {
        self.stage(1, |inner| inner.move_custom_table_styles(placement))
    }

    /// Set one slide's `SlidePersistAtom.fNonOutlineData` flag ([MS-PPT]
    /// 2.4.14.5; record 0357a). It does not hide the slide; that is
    /// [`Self::set_slide_hidden`].
    ///
    /// # Errors
    ///
    /// [`slide_order::Refusal::SlideNotFound`] for a position outside the
    /// slide list.
    pub fn set_slide_non_outline_data(&mut self, position: Position, present: bool) -> Result<()> {
        self.stage(1, |inner| {
            inner.set_slide_non_outline_data(position, present)
        })
    }

    /// Publish this edit as a new immutable package.
    pub fn commit(self) -> Result<Commit> {
        self.commit_with(slide_order::Transaction::commit)
    }

    /// Publish this edit while reporting deterministic, content-free phase
    /// boundaries of the native publication to `observer`.
    ///
    /// This is the opt-in diagnostic equivalent of [`Self::commit`]; it keeps
    /// every validation and ownership rule of the ordinary commit.
    #[cfg(feature = "performance-diagnostics")]
    pub fn commit_profiled(
        self,
        observer: impl FnMut(slide_order::DiagnosticEvent),
    ) -> Result<Commit> {
        self.commit_with(|inner| inner.commit_profiled(observer))
    }

    fn commit_with(
        self,
        native: impl FnOnce(slide_order::Transaction) -> slide_order::Result<slide_order::Commit>,
    ) -> Result<Commit> {
        ensure_source_version(&self.package, self.expected_source)?;
        let verified = self.inner()?.verified_shapes();
        let inner = self.inner.into_inner().ok_or_else(|| {
            Error::Package(PackageError::Corrupted(
                "PPT edit native state was not initialized".into(),
            ))
        })?;
        let native_commit = native(inner);
        ensure_source_version(&self.package, self.expected_source)?;
        let native_commit = native_commit?;
        let operation_count = self.operation_count;
        let (snapshot, native_patch) = native_commit.into_parts();
        let patch = Patch::from_native(native_patch, operation_count);
        let source = self.package.clone();
        let package = if patch.is_noop() {
            Ok(source.clone())
        } else {
            package_from_bytes(&source, snapshot.bytes()).inspect(|package| {
                seed_verified_shapes(&source, package, snapshot.bytes(), verified);
            })
        };
        // This is the final source fence, after the target has been fully
        // parsed.  The source owner remains borrowed through `source` until
        // this check completes.
        ensure_source_version(&source, self.expected_source)?;
        let package = package?;
        let diagnostics = Diagnostics::new(!patch.is_noop(), operation_count);
        Ok(Commit {
            package,
            patch,
            diagnostics,
        })
    }

    fn stage<T>(
        &mut self,
        charge: usize,
        operation: impl FnOnce(&mut slide_order::Transaction) -> slide_order::Result<T>,
    ) -> Result<T> {
        ensure_source_version(&self.package, self.expected_source)?;
        self.reserve_operations(charge)?;
        let mut candidate = self.inner()?.clone();
        let result = operation(&mut candidate);
        // Check the retained source even when the native operation failed.
        // A source change must never be hidden by a later semantic refusal.
        ensure_source_version(&self.package, self.expected_source)?;
        let result = result.map_err(Error::from)?;
        *self.inner_mut()? = candidate;
        self.operation_count = self
            .operation_count
            .checked_add(charge)
            .ok_or(Error::Limit {
                kind: LimitKind::Operations,
                observed: usize::MAX,
                maximum: self.options.max_operations,
            })?;
        Ok(result)
    }

    fn reserve_operations(&self, charge: usize) -> Result<()> {
        let observed = self
            .operation_count
            .checked_add(charge)
            .ok_or(Error::Limit {
                kind: LimitKind::Operations,
                observed: usize::MAX,
                maximum: self.options.max_operations,
            })?;
        if observed > self.options.max_operations {
            return Err(Error::Limit {
                kind: LimitKind::Operations,
                observed,
                maximum: self.options.max_operations,
            });
        }
        Ok(())
    }
}

impl Package {
    /// Start the canonical primary PPT transaction under the default profile.
    pub fn edit(&self) -> Result<Edit> {
        Edit::from_package_with_options(self.clone(), EditOptions::default())
    }

    /// Start the canonical primary PPT transaction under a stricter profile.
    pub fn edit_with_options(&self, options: EditOptions) -> Result<Edit> {
        Edit::from_package_with_options(self.clone(), options)
    }

    /// Apply a primary exact-source patch without mutating this package.
    ///
    /// The commit carries the published package, a copy of `patch` and its
    /// diagnostics; an exact no-op patch publishes this package unchanged.
    pub fn apply(&self, patch: &Patch) -> Result<Commit> {
        let package = self.apply_package(patch)?;
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics::new(!patch.is_noop(), patch.operation_count),
        })
    }

    /// Replay a durable patch (see [`Patch::to_durable`]) on this package.
    ///
    /// Each operation checks its own preconditions (the exact artifact
    /// SHA-256 of structural operations, the before value of property
    /// operations). The commit's patch is exact (`apply`, `inverse`) and
    /// counts the durable operations; it keeps no semantic operations of its
    /// own. An empty durable patch publishes this package unchanged.
    ///
    /// # Errors
    ///
    /// [`Error::Native`] for another format, malformed operations or failed
    /// preconditions, [`Error::Signed`] for a changed replay on a signed
    /// package, and [`Error::Package`] for an encrypted package.
    pub fn apply_durable<Mode>(&self, patch: &litchi_core::patch::Patch<Mode>) -> Result<Commit> {
        if self.is_encrypted() {
            return Err(Error::Package(PackageError::EncryptionPolicy));
        }
        let expected = self.source_version().map_err(Error::Package)?;
        let bytes = self.to_encrypted_bytes().map_err(Error::Package)?;
        ensure_source_version(self, expected)?;
        let replay = (|| {
            let source = slide_order::Snapshot::from_bytes_with_options(bytes, open_options(self))?;
            let target = source.apply_durable(patch)?;
            Ok::<_, slide_order::Error>((source, target))
        })();
        ensure_source_version(self, expected)?;
        let (source, target) = replay?;
        let operation_count = patch.operations().len();
        let changed = source.bytes() != target.bytes();
        let package = if changed {
            package_from_bytes(self, target.bytes())?
        } else {
            self.clone()
        };
        ensure_source_version(self, expected)?;
        let mut patch =
            Patch::from_native(slide_order::Patch::exact(source, target), operation_count);
        patch.replayed = true;
        Ok(Commit {
            package,
            patch,
            diagnostics: Diagnostics::new(changed, operation_count),
        })
    }

    /// Plan two patches made from this exact package against each other
    /// without applying either.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when either patch starts from another
    /// package.
    pub fn plan_three_way(&self, left: &Patch, right: &Patch) -> Result<ThreeWayPlan> {
        if self.is_encrypted() {
            return Err(Error::Package(PackageError::EncryptionPolicy));
        }
        let expected = self.source_version().map_err(Error::Package)?;
        let bytes = self.to_encrypted_bytes().map_err(Error::Package)?;
        ensure_source_version(self, expected)?;
        if left.before() != bytes || right.before() != bytes {
            return Err(Error::PatchConflict);
        }
        let native = slide_order::ThreeWayPlan::new(&left.inner, &right.inner);
        Ok(ThreeWayPlan {
            conflicts: native
                .conflicts()
                .iter()
                .map(|conflict| Conflict {
                    target: conflict.target().to_owned(),
                })
                .collect(),
        })
    }

    fn apply_package(&self, patch: &Patch) -> Result<Self> {
        if self.is_encrypted() {
            return Err(Error::Package(PackageError::EncryptionPolicy));
        }
        let expected = self.source_version().map_err(Error::Package)?;
        let bytes = self.to_encrypted_bytes().map_err(Error::Package)?;
        ensure_source_version(self, expected)?;
        if bytes != patch.before() {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            ensure_source_version(self, expected)?;
            return Ok(self.clone());
        }
        let snapshot = slide_order::Snapshot::from_bytes_with_options(bytes, open_options(self));
        ensure_source_version(self, expected)?;
        let target = patch.inner.apply(&snapshot?);
        ensure_source_version(self, expected)?;
        let target = target?;
        let package = package_from_bytes(self, target.bytes());
        ensure_source_version(self, expected)?;
        package
    }

    /// Prepare a transfer plan from `donor` for insertion into this package.
    ///
    /// Both packages are fenced before and after native planning.  The
    /// returned plan is immutable and can be supplied to
    /// [`Edit::insert_transfer`].
    pub fn plan_transfer_from(&self, donor: &Package, position: Position) -> Result<TransferPlan> {
        if self.is_encrypted() || donor.is_encrypted() {
            return Err(Error::Package(PackageError::EncryptionPolicy));
        }
        let expected = self.source_version().map_err(Error::Package)?;
        let donor_expected = donor.source_version().map_err(Error::Package)?;
        let target_bytes = self.to_encrypted_bytes().map_err(Error::Package)?;
        let donor_bytes = donor.to_encrypted_bytes().map_err(Error::Package)?;
        ensure_source_version(self, expected)?;
        ensure_source_version(donor, donor_expected)?;
        let plan = (|| {
            let target =
                slide_order::Snapshot::from_bytes_with_options(target_bytes, open_options(self))?;
            let donor_snapshot =
                slide_order::Snapshot::from_bytes_with_options(donor_bytes, open_options(donor))?;
            target.plan_transfer_from(&donor_snapshot, position)
        })();
        ensure_source_version(self, expected)?;
        ensure_source_version(donor, donor_expected)?;
        Ok(plan?)
    }
}

/// Parse exactly one record from `atom` (the custom table-style verbs).
fn custom_table_styles_record(atom: &[u8]) -> Result<crate::Record> {
    let invalid = || {
        Error::Native(slide_order::Error::Refused(
            slide_order::Refusal::InvalidElementValue {
                element: "custom table styles",
                reason: "the bytes are not exactly one record",
            },
        ))
    };
    let (record, consumed) = crate::Record::parse(atom, 0, &crate::SourceLimits::default())
        .map_err(|_error| invalid())?;
    if consumed != atom.len() {
        return Err(invalid());
    }
    Ok(record)
}

fn encode_atom(record: &crate::Record) -> Result<Vec<u8>> {
    crate::document_structure::encode_record(record, false).map_err(Error::Package)
}

fn open_options(source: &Package) -> OpenOptions<'static> {
    let options = OpenOptions::new()
        .source_limits(source.source_limits())
        .leniency(source.leniency());
    #[cfg(feature = "vba-inspection")]
    let options = options.vba_project_limits(source.vba_project_limits());
    options
}

/// Hand the shapes a batch text publication proved for its output to the
/// published package's slide memo, so the next chunked edit resolves its
/// targets without parsing the slide again (record 0264).
///
/// The proof was made against a default-policy reopen of exactly the
/// published bytes, from a source parse made by the source package. It
/// transfers only when the published package was opened from the same
/// allocation under the same (default) policies.
fn seed_verified_shapes(
    source: &Package,
    published: &Package,
    bytes: &[u8],
    verified: Option<slide_order::VerifiedSlideShapes>,
) {
    let Some(verified) = verified else {
        return;
    };
    if !std::ptr::eq(verified.bytes.as_ref(), bytes)
        || source.source_limits() != crate::SourceLimits::default()
        || source.leniency() != crate::leniency::Leniency::default()
    {
        return;
    }
    published.seed_slide_shapes(verified.persist_id, verified.shapes);
}

fn package_from_bytes(source: &Package, bytes: &[u8]) -> Result<Package> {
    Package::from_bytes_with_options(bytes, open_options(source)).map_err(Error::Package)
}

fn ensure_source_version(package: &Package, expected: SourceVersion) -> Result<()> {
    let observed = package.source_version().map_err(Error::Package)?;
    if observed != expected {
        return Err(Error::Package(PackageError::Ole(
            litchi_cfb::OleError::SourceChanged { expected, observed },
        )));
    }
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "the lazy-entry regression is a direct state assertion"
)]
mod tests {
    use super::{Edit, EditOptions, Package};

    #[test]
    fn primary_edit_construction_and_drop_do_not_open_native_state() {
        let package = Package::new().unwrap();
        let edit = Edit::from_package_with_options(package, EditOptions::new()).unwrap();
        assert!(edit.inner.get().is_none());
        drop(edit);
    }
}
