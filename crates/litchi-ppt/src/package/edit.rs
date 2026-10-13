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
pub use crate::external_media::Playback;
use crate::slide_order;
pub use crate::slide_order::{
    ShapeTextReplacement, SlideAdvance, SlideTransitionVisual, TransferPlan,
};
pub use crate::text_edit::Target;
pub use litchi_core::Position;
use litchi_core::SourceVersion;
use std::fmt;

/// The finite operation dimension of the primary PPT edit.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Number of successful staged edit operations.
    Operations,
}

/// The default and hard bound for one primary PPT edit.
pub const DEFAULT_MAX_OPERATIONS: usize = 4096;
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
    pub const fn get_max_operations(self) -> usize {
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
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Package(error) => Some(error),
            Self::Native(error) => Some(error),
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
            Self::Package(_) | Self::Native(_) | Self::PatchConflict | Self::Signed => None,
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
            Self::Package(_) | Self::Native(_) | Self::PatchConflict | Self::Signed => None,
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
        Self::from_native(self.inner.inverse(), self.operation_count)
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
    inner: slide_order::Transaction,
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
        let bytes = package.to_encrypted_bytes().map_err(Error::Package)?;
        ensure_source_version(&package, expected_source)?;
        let snapshot =
            slide_order::Snapshot::from_bytes_with_options(bytes, open_options(&package));
        ensure_source_version(&package, expected_source)?;
        let inner = snapshot?.edit();
        ensure_source_version(&package, expected_source)?;
        // The package's own projection was parsed from exactly these bytes;
        // target resolution reuses its parsed slides (record 0260).
        let inner = inner?.with_opened_projection(package.presentation().map_err(Error::Package)?);
        Ok(Self {
            package,
            inner,
            options,
            operation_count: 0,
            expected_source,
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
    /// including empty chunks, and at most 4096 replacements across the whole
    /// call. Targets must be globally distinct. The operation count charged
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

    /// Publish this edit as a new immutable package.
    pub fn commit(self) -> Result<Commit> {
        ensure_source_version(&self.package, self.expected_source)?;
        let native_commit = self.inner.commit();
        ensure_source_version(&self.package, self.expected_source)?;
        let native_commit = native_commit?;
        let operation_count = self.operation_count;
        let (snapshot, native_patch) = native_commit.into_parts();
        let patch = Patch::from_native(native_patch, operation_count);
        let source = self.package.clone();
        let package = if patch.is_noop() {
            Ok(source.clone())
        } else {
            package_from_bytes(&source, snapshot.bytes())
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
        let mut candidate = self.inner.clone();
        let result = operation(&mut candidate);
        // Check the retained source even when the native operation failed.
        // A source change must never be hidden by a later semantic refusal.
        ensure_source_version(&self.package, self.expected_source)?;
        let result = result.map_err(Error::from)?;
        self.inner = candidate;
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
    pub fn apply(&self, patch: &Patch) -> Result<Self> {
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

fn open_options(source: &Package) -> OpenOptions<'static> {
    OpenOptions::new()
        .source_limits(source.source_limits())
        .leniency(source.leniency())
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
