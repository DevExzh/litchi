//! Canonical primary DOC package transaction.
//!
//! [`Package`] is the immutable owner used by this module.  The transaction
//! delegates its checked semantic operations to [`crate::body_text`], which
//! remains the source-preserving owner for ordinary DOC text and its bounded
//! dependency closures.  The package boundary adds the canonical
//! package-owned publication shape without introducing a second edit engine.
//!
//! Opening an edit retains only the immutable package and its effective
//! options. The first body operation then copies and revalidates the retained
//! source through the body owner. Changed candidates transfer that owner's
//! already validated CFB and document projection into the package owner; exact
//! no-ops retain the original package. The body engine's exact-source,
//! protection, and full-candidate validation gates remain authoritative at
//! every publication boundary.

use super::{Error as PackageError, OpenOptions, Package, model::EditProfile};
use crate::body_text;
use crate::parts::dofr::Patch as DofrPatch;
use crate::parts::protection::ProtectionPolicy;
use crate::parts::saved_selection::{Edit as SavedSelectionEdit, Patch as SavedSelectionPatch};
use crate::tracked_revision::Limits;
use litchi_core::Position;
use std::fmt;

/// A canonical primary transaction limit dimension.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Number of semantic or auxiliary operations staged by one edit.
    Operations,
    /// UTF-16 units in one replacement payload.
    ReplacementUnits,
    /// Aggregate UTF-16 replacement units in one transaction.
    TotalReplacementUnits,
    /// Capacity retained for one validated rendered candidate.
    RetainedRenderBytes,
}

/// Caller-selected bounds and protection policy for a primary package edit.
///
/// Physical package limits remain attached to the source [`Package`] and are
/// always intersected with the body profile here.  The body revision and
/// transaction profiles are likewise intersected by the underlying checked
/// engine before any candidate is staged.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct EditOptions {
    revision_limits: Limits,
    transaction_limits: body_text::TransactionLimits,
    protection_policy: ProtectionPolicy,
}

impl EditOptions {
    /// Create the default bounded edit profile.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Select strict tracked-revision parsing limits.
    #[must_use]
    pub const fn revision_limits(mut self, limits: Limits) -> Self {
        self.revision_limits = limits;
        self
    }

    /// Select finite staged-operation and replacement limits.
    #[must_use]
    pub const fn transaction_limits(mut self, limits: body_text::TransactionLimits) -> Self {
        self.transaction_limits = limits;
        self
    }

    /// Select the explicit policy for changed protected-source publication.
    #[must_use]
    pub fn protection_policy(mut self, policy: ProtectionPolicy) -> Self {
        self.protection_policy = policy;
        self
    }

    /// Return the tracked-revision limits.
    #[must_use]
    pub const fn revision_limits_value(&self) -> Limits {
        self.revision_limits
    }

    /// Return the transaction limits.
    #[must_use]
    pub const fn transaction_limits_value(&self) -> body_text::TransactionLimits {
        self.transaction_limits
    }

    /// Return the protected-publication policy.
    #[must_use]
    pub fn protection_policy_value(&self) -> ProtectionPolicy {
        self.protection_policy.clone()
    }
}

impl EditProfile {
    fn from_snapshot(snapshot: &body_text::Snapshot) -> Self {
        Self {
            revision_limits: snapshot.revision_limits(),
            transaction_limits: snapshot.transaction_limits(),
            protection_policy: snapshot.protection_policy(),
        }
    }

    fn from_patch(patch: &body_text::Patch) -> Self {
        Self::from_snapshot(patch.before())
    }

    fn options(&self) -> EditOptions {
        EditOptions {
            revision_limits: self.revision_limits,
            transaction_limits: self.transaction_limits,
            protection_policy: self.protection_policy.clone(),
        }
    }

    fn constrain_options(&self, options: EditOptions) -> EditOptions {
        EditOptions {
            revision_limits: body_text::intersect_revision_limits(
                self.revision_limits,
                options.revision_limits,
            ),
            transaction_limits: self
                .transaction_limits
                .intersect(options.transaction_limits),
            // An explicit caller policy is a capability decision. Numeric
            // bounds are retained by intersection, while protection may be
            // deliberately selected for this edit.
            protection_policy: options.protection_policy,
        }
    }
}

/// Typed error at the package transaction boundary.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// Physical package ingress, publication, or source-fence failure.
    Package(PackageError),
    /// A caller-selected body transaction bound was exceeded.
    Limit {
        /// Dimension whose bound was exceeded.
        kind: LimitKind,
        /// Observed operation or replacement amount.
        observed: usize,
        /// Selected maximum.
        maximum: usize,
    },
    /// Checked body transaction staging or publication failure.
    Body(body_text::Error),
}

impl From<PackageError> for Error {
    fn from(error: PackageError) -> Self {
        Self::Package(error)
    }
}

impl From<body_text::Error> for Error {
    fn from(error: body_text::Error) -> Self {
        match error {
            body_text::Error::Refused(body_text::Refusal::OperationLimit { observed, limit }) => {
                Self::Limit {
                    kind: LimitKind::Operations,
                    observed,
                    maximum: limit,
                }
            },
            body_text::Error::Refused(body_text::Refusal::ReplacementLimit { observed, limit }) => {
                Self::Limit {
                    kind: LimitKind::TotalReplacementUnits,
                    observed,
                    maximum: limit,
                }
            },
            error => Self::Body(error),
        }
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package(error) => error.fmt(formatter),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "DOC edit limit {kind:?}: {observed} exceeds {maximum}"
            ),
            Self::Body(error) => error.fmt(formatter),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Package(error) => Some(error),
            Self::Limit { .. } => None,
            Self::Body(error) => Some(error),
        }
    }
}

impl Error {
    /// Return the exceeded limit dimension, when this is a limit failure.
    #[must_use]
    pub const fn limit_kind(&self) -> Option<LimitKind> {
        match self {
            Self::Limit { kind, .. } => Some(*kind),
            Self::Package(_) | Self::Body(_) => None,
        }
    }

    /// Return the observed amount and maximum, when this is a limit failure.
    #[must_use]
    pub const fn limit_values(&self) -> Option<(usize, usize)> {
        match self {
            Self::Limit {
                observed, maximum, ..
            } => Some((*observed, *maximum)),
            Self::Package(_) | Self::Body(_) => None,
        }
    }
}

/// Result type for the canonical package transaction.
pub type Result<T> = std::result::Result<T, Error>;

/// Deferred-open state for an ordinary package edit.
///
/// This is public only as the default parameter of [`Edit`]. Keeping the
/// observer state out of this value preserves the ordinary edit's `Send` and
/// `Sync` auto traits when diagnostics are enabled.
#[doc(hidden)]
#[derive(Clone, Copy, Debug, Default)]
pub struct OrdinaryEditState;

trait DeferredOpenState {
    fn open(&mut self, package: &Package, options: &EditOptions) -> Result<body_text::Edit>;
}

impl DeferredOpenState for OrdinaryEditState {
    fn open(&mut self, package: &Package, options: &EditOptions) -> Result<body_text::Edit> {
        let source = package.body_snapshot(options)?;
        source.edit().map_err(Error::from)
    }
}

#[cfg(feature = "performance-diagnostics")]
/// Deferred observer state for a profiled package edit.
#[doc(hidden)]
pub struct ProfiledOpen<'a> {
    observer: Box<dyn FnMut(body_text::DiagnosticEvent) + 'a>,
    cfb_observer: Box<dyn FnMut(body_text::CfbParseEvent) + 'a>,
}

#[cfg(feature = "performance-diagnostics")]
impl<'a> DeferredOpenState for ProfiledOpen<'a> {
    fn open(&mut self, package: &Package, options: &EditOptions) -> Result<body_text::Edit> {
        let snapshot = body_text::Snapshot::open_bounded_profiled_with_cfb_observer(
            package.to_bytes().map_err(Error::Package)?,
            options.revision_limits,
            options.transaction_limits,
            package.source_limits(),
            options.protection_policy.clone(),
            &mut *self.observer,
            &mut *self.cfb_observer,
        )
        .map_err(Error::from)?;
        snapshot
            .edit_with_options(options.clone())
            .map_err(Error::from)
    }
}

/// Content-free facts about one published package transaction.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    operation_count: usize,
    auxiliary_changed: bool,
}

impl Diagnostics {
    fn from_patch(patch: &body_text::Patch) -> Self {
        Self {
            changed: !patch.is_noop(),
            operation_count: patch.changes().len(),
            auxiliary_changed: patch.has_auxiliary_changes(),
        }
    }

    /// Whether publication changed the exact package bytes.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of semantic body changes represented by the patch.
    #[must_use]
    pub const fn operation_count(self) -> usize {
        self.operation_count
    }

    /// Whether the commit included an inert auxiliary-table change.
    #[must_use]
    pub const fn has_auxiliary_changes(self) -> bool {
        self.auxiliary_changed
    }
}

/// Reversible exact-source package patch produced by a primary [`Edit`].
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    pub(super) inner: body_text::Patch,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("source_fingerprint", &self.source_fingerprint())
            .field("target_fingerprint", &self.target_fingerprint())
            .field("operation_count", &self.inner.changes().len())
            .field("auxiliary_changed", &self.has_auxiliary_changes())
            .finish()
    }
}

impl Patch {
    fn from_body(inner: body_text::Patch) -> Self {
        Self { inner }
    }

    /// Numbered semantic changes in staging order.
    #[must_use]
    pub fn changes(&self) -> impl ExactSizeIterator<Item = body_text::ChangeRef<'_>> {
        self.inner.changes()
    }

    /// Effective body transaction limits retained by this edit.
    #[must_use]
    pub const fn transaction_limits_value(&self) -> body_text::TransactionLimits {
        self.inner.before().transaction_limits()
    }

    /// Whether the patch includes a checked inert auxiliary-table change.
    #[must_use]
    pub const fn has_auxiliary_changes(&self) -> bool {
        self.inner.has_auxiliary_changes()
    }

    /// Fast diagnostic fingerprint of the exact source bytes.
    #[must_use]
    pub fn source_fingerprint(&self) -> u64 {
        self.inner.source_fingerprint()
    }

    /// Fast diagnostic fingerprint of the exact target bytes.
    #[must_use]
    pub fn target_fingerprint(&self) -> u64 {
        self.inner.target_fingerprint()
    }

    /// Exact source bytes required by this patch.
    ///
    /// This is an explicit raw-boundary accessor for package owners that need
    /// to compare a patch against an already retained source.  Fingerprints
    /// remain diagnostics only; callers must use these bytes for authority.
    #[must_use]
    pub fn before_bytes(&self) -> &[u8] {
        self.inner.before().bytes()
    }

    /// Exact target bytes produced by this patch.
    #[must_use]
    pub fn after_bytes(&self) -> &[u8] {
        self.inner.after().bytes()
    }

    /// Whether applying this patch preserves the exact source bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.inner.is_noop()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self::from_body(self.inner.inverse())
    }

    /// Convert the checked body patch to the shared durable patch envelope.
    pub fn to_durable(
        &self,
        limits: litchi_core::patch::PatchLimits,
    ) -> std::result::Result<
        litchi_core::patch::Patch<litchi_core::patch::Reversible>,
        litchi_core::patch::PatchError,
    > {
        self.inner.to_durable(limits)
    }
}

/// A package-owned commit carrying the canonical publication triad: the
/// published package, its reversible exact-source patch and content-free
/// diagnostics.
pub type Commit = litchi_core::edit::Commit<Package, Patch, Diagnostics>;

fn commit_from_body(source: Package, body_commit: body_text::PackageCommit) -> Result<Commit> {
    let body_text::PackageCommit {
        snapshot,
        patch: body_patch,
        validated,
    } = body_commit;
    let patch = Patch::from_body(body_patch);
    let diagnostics = Diagnostics::from_patch(&patch.inner);
    let package = match validated {
        None => source,
        // The commit validated the candidate's document strictly under the
        // snapshot's source limits, whatever leniency the source retains.
        Some(validated) => Package::from_validated_ole(
            validated.ole,
            validated.document,
            snapshot.source_limits(),
            source.leniency(),
            true,
        )
        .map_err(Error::Package)?,
    };
    Ok(Commit::new(
        package.with_edit_profile(EditProfile::from_snapshot(&snapshot)),
        patch,
        diagnostics,
    ))
}
/// Primary immutable DOC package transaction.
pub struct Edit<'a, S = OrdinaryEditState> {
    package: Package,
    options: EditOptions,
    inner: Option<body_text::Edit>,
    state: S,
    _profile: std::marker::PhantomData<&'a ()>,
}

impl<'a, S> fmt::Debug for Edit<'a, S> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("package", &self.package)
            .field(
                "retained_render_bytes",
                &self
                    .inner
                    .as_ref()
                    .and_then(body_text::Edit::retained_render_bytes),
            )
            .finish_non_exhaustive()
    }
}

impl Edit<'static, OrdinaryEditState> {
    pub(crate) fn from_package_with_options(
        package: Package,
        options: EditOptions,
    ) -> Result<Self> {
        if package.document().fib().is_encrypted() {
            return Err(Error::Body(body_text::Error::Refused(
                body_text::Refusal::Encrypted,
            )));
        }
        Ok(Self {
            package,
            options,
            inner: None,
            state: OrdinaryEditState,
            _profile: std::marker::PhantomData,
        })
    }

    /// A primary edit of `package` over a body edit already staged from its
    /// body snapshot (a composition, a merge plan or a durable replay).
    pub(crate) fn from_body(package: Package, inner: body_text::Edit) -> Self {
        Self {
            options: package.retained_edit_options(),
            package,
            inner: Some(inner),
            state: OrdinaryEditState,
            _profile: std::marker::PhantomData,
        }
    }
}

#[cfg(feature = "performance-diagnostics")]
impl<'a> Edit<'a, ProfiledOpen<'a>> {
    pub(crate) fn from_profiled(
        package: Package,
        options: EditOptions,
        observer: impl FnMut(body_text::DiagnosticEvent) + 'a,
        cfb_observer: impl FnMut(body_text::CfbParseEvent) + 'a,
    ) -> Self {
        Self {
            package,
            options,
            inner: None,
            state: ProfiledOpen {
                observer: Box::new(observer),
                cfb_observer: Box::new(cfb_observer),
            },
            _profile: std::marker::PhantomData,
        }
    }

    /// Finish the deferred profiled open and return the ordinary transferable
    /// edit after its observers have been consumed.
    #[must_use = "the profiled edit must be retained or converted"]
    pub fn into_unprofiled(mut self) -> Result<Edit<'static>> {
        self.inner_mut()?;
        let Edit {
            package,
            options,
            inner,
            state: _,
            _profile: _,
        } = self;
        Ok(Edit {
            package,
            options,
            inner,
            state: OrdinaryEditState,
            _profile: std::marker::PhantomData,
        })
    }
}

#[allow(private_bounds)]
impl<'a, S: DeferredOpenState> Edit<'a, S> {
    fn inner_mut(&mut self) -> Result<&mut body_text::Edit> {
        if self.inner.is_none() {
            self.inner = Some(self.state.open(&self.package, &self.options)?);
        }
        self.inner.as_mut().ok_or_else(|| {
            Error::Package(PackageError::InvalidFormat(
                "DOC edit was not initialized".into(),
            ))
        })
    }

    fn take_inner(&mut self) -> Result<body_text::Edit> {
        self.inner_mut()?;
        self.inner.take().ok_or_else(|| {
            Error::Package(PackageError::InvalidFormat(
                "DOC edit was not initialized".into(),
            ))
        })
    }

    /// Immutable package snapshot that authorizes this transaction.
    #[must_use]
    pub const fn source(&self) -> &Package {
        &self.package
    }

    /// Exact source bytes retained by this staged edit.
    pub fn source_bytes(&mut self) -> Result<&[u8]> {
        Ok(self.inner_mut()?.source_bytes())
    }

    /// Borrowed semantic changes staged so far, including authored text.
    pub fn changes(&mut self) -> Result<impl ExactSizeIterator<Item = body_text::ChangeRef<'_>>> {
        Ok(self.inner_mut()?.changes())
    }

    /// Allocation capacity held by the validated candidate, if retained.
    #[must_use]
    pub fn retained_render_bytes(&self) -> Option<usize> {
        self.inner
            .as_ref()
            .and_then(body_text::Edit::retained_render_bytes)
    }

    /// Release the retained candidate without changing staged semantics.
    pub fn release_retained_render(&mut self) -> Result<()> {
        self.inner_mut()?.release_retained_render();
        Ok(())
    }

    /// Apply a checked `Selsf` patch to this transaction.
    pub fn apply_saved_selection_patch(&mut self, patch: &SavedSelectionPatch) -> Result<bool> {
        self.inner_mut()?
            .apply_saved_selection_patch(patch)
            .map_err(Error::from)
    }

    /// Edit the optional `Selsf` record through its typed transaction.
    pub fn edit_saved_selection(
        &mut self,
        edit: impl FnOnce(&mut SavedSelectionEdit) -> crate::package::Result<()>,
    ) -> Result<bool> {
        self.inner_mut()?
            .edit_saved_selection(edit)
            .map_err(Error::from)
    }

    /// Apply a checked `RgDofr` patch to this transaction.
    pub fn apply_dofr_patch(&mut self, patch: &DofrPatch) -> Result<bool> {
        self.inner_mut()?
            .apply_dofr_patch(patch)
            .map_err(Error::from)
    }

    /// Replace one complete checked `RgDofr` record.
    pub fn replace_dofr_record(&mut self, index: usize, replacement: &[u8]) -> Result<bool> {
        self.inner_mut()?
            .replace_dofr_record(index, replacement)
            .map_err(Error::from)
    }

    /// Replace one ordinary main-story paragraph.
    pub fn replace_paragraph(&mut self, position: Position, replacement: &str) -> Result<()> {
        self.inner_mut()?
            .replace_paragraph(position, replacement)
            .map_err(Error::from)
    }

    /// Replace one checked body, table-cell, or field target.
    pub fn replace_text(&mut self, target: body_text::TextTarget, replacement: &str) -> Result<()> {
        self.inner_mut()?
            .replace_text(target, replacement)
            .map_err(Error::from)
    }

    /// Set direct bold on one ordinary paragraph.
    pub fn set_paragraph_bold(&mut self, position: Position, enabled: bool) -> Result<()> {
        self.inner_mut()?
            .set_paragraph_bold(position, enabled)
            .map_err(Error::from)
    }

    /// Set one supported direct character property.
    pub fn set_character_property(
        &mut self,
        target: body_text::TextTarget,
        property: body_text::CharacterProperty,
        enabled: bool,
    ) -> Result<()> {
        self.inner_mut()?
            .set_character_property(target, property, enabled)
            .map_err(Error::from)
    }

    /// Set one valued direct character property (size, color or font).
    pub fn set_character_value(
        &mut self,
        target: body_text::TextTarget,
        value: body_text::CharacterValue<'_>,
    ) -> Result<()> {
        self.inner_mut()?
            .set_character_value(target, value)
            .map_err(Error::from)
    }

    /// Apply paragraph style `istd` to one ordinary main-story paragraph
    /// (record 0324).
    pub fn set_paragraph_style(&mut self, position: Position, istd: u16) -> Result<()> {
        self.inner_mut()?
            .set_paragraph_style(position, istd)
            .map_err(Error::from)
    }

    /// The effective tab stops of one ordinary main-story paragraph.
    pub fn paragraph_tabs(
        &mut self,
        position: Position,
    ) -> Result<Vec<crate::parts::pap::TabStop>> {
        self.inner_mut()?
            .paragraph_tabs(position)
            .map_err(Error::from)
    }

    /// Set the tab stops of one ordinary main-story paragraph.
    pub fn set_paragraph_tabs(
        &mut self,
        position: Position,
        stops: &[crate::parts::pap::TabStop],
    ) -> Result<()> {
        self.inner_mut()?
            .set_paragraph_tabs(position, stops)
            .map_err(Error::from)
    }

    /// Insert a `HYPERLINK` field into one ordinary main-story paragraph,
    /// before UTF-16 offset `offset` of its text.
    pub fn insert_hyperlink(
        &mut self,
        position: Position,
        offset: u32,
        link: &body_text::HyperlinkInsert,
    ) -> Result<()> {
        self.inner_mut()?
            .insert_hyperlink(position, offset, link)
            .map_err(Error::from)
    }

    /// Remove the whole main-story field that begins at `start_cp`.
    pub fn remove_field(&mut self, start_cp: u32) -> Result<()> {
        self.inner_mut()?
            .remove_field(start_cp)
            .map_err(Error::from)
    }

    /// The standard bookmarks, as (name, start CP, end CP) in start order.
    pub fn bookmarks(&mut self) -> Result<Vec<(String, u32, u32)>> {
        self.inner_mut()?.bookmarks().map_err(Error::from)
    }

    /// Add a standard bookmark over a UTF-16 range of one ordinary
    /// main-story paragraph.
    pub fn add_bookmark(
        &mut self,
        name: &str,
        position: Position,
        start: u32,
        end: u32,
    ) -> Result<()> {
        self.inner_mut()?
            .add_bookmark(name, position, start, end)
            .map_err(Error::from)
    }

    /// Rename the standard bookmark named exactly `from` to `to`.
    pub fn rename_bookmark(&mut self, from: &str, to: &str) -> Result<()> {
        self.inner_mut()?
            .rename_bookmark(from, to)
            .map_err(Error::from)
    }

    /// Remove the standard bookmark named exactly `name`.
    pub fn remove_bookmark(&mut self, name: &str) -> Result<()> {
        self.inner_mut()?.remove_bookmark(name).map_err(Error::from)
    }

    /// Set page-setup values on one section.
    pub fn set_section_values(
        &mut self,
        section: Position,
        values: &[body_text::SectionValue],
    ) -> Result<()> {
        self.inner_mut()?
            .set_section_values(section, values)
            .map_err(Error::from)
    }

    /// Insert one ordinary main-story paragraph before `position`, or after
    /// the last one when `position` is the paragraph count.
    pub fn insert_paragraph(&mut self, position: Position, text: &str) -> Result<()> {
        self.inner_mut()?
            .insert_paragraph(position, text)
            .map_err(Error::from)
    }

    /// Remove one ordinary main-story paragraph.
    pub fn remove_paragraph(&mut self, position: Position) -> Result<()> {
        self.inner_mut()?
            .remove_paragraph(position)
            .map_err(Error::from)
    }

    /// Move one ordinary main-story paragraph so that it ends at `to`,
    /// counted after it is detached.
    pub fn move_paragraph(&mut self, from: Position, to: Position) -> Result<()> {
        self.inner_mut()?
            .move_paragraph(from, to)
            .map_err(Error::from)
    }

    /// Apply a dependency-checked text transfer.
    pub fn apply_transfer(&mut self, plan: &body_text::TransferPlan) -> Result<()> {
        self.inner_mut()?.apply_transfer(plan).map_err(Error::from)
    }

    /// Apply a dependency-closed embedded-object transfer.
    pub fn apply_embedded_transfer(
        &mut self,
        plan: &body_text::EmbeddedTransferPlan,
    ) -> Result<()> {
        self.inner_mut()?
            .apply_embedded_transfer(plan)
            .map_err(Error::from)
    }

    /// Apply a checked inline or floating picture transfer.
    pub fn apply_picture_transfer(&mut self, plan: &body_text::PictureTransferPlan) -> Result<()> {
        self.inner_mut()?
            .apply_picture_transfer(plan)
            .map_err(Error::from)
    }

    /// Add one inert embedded object through its dependency owner.
    pub fn add_embedded_object(
        &mut self,
        options: crate::embedded_object::WriteOptions,
    ) -> Result<()> {
        self.inner_mut()?
            .add_embedded_object(options)
            .map_err(Error::from)
    }

    /// Remove one managed embedded object and its dependency closure.
    pub fn remove_embedded_object(&mut self, storage_id: u32) -> Result<()> {
        self.inner_mut()?
            .remove_embedded_object(storage_id)
            .map_err(Error::from)
    }

    /// Accept or reject one safe revision mark.
    pub fn dispose_revision(
        &mut self,
        position: Position,
        disposition: body_text::RevisionDisposition,
    ) -> Result<()> {
        self.inner_mut()?
            .dispose_revision(position, disposition)
            .map_err(Error::from)
    }

    /// Change the passive display-as-icon bit of one managed object.
    pub fn set_embedded_display_as_icon(&mut self, storage_id: u32, enabled: bool) -> Result<()> {
        self.inner_mut()?
            .set_embedded_display_as_icon(storage_id, enabled)
            .map_err(Error::from)
    }

    /// Publish the checked package commit and canonical triad.
    pub fn commit(mut self) -> Result<Commit> {
        let body_commit = self
            .take_inner()?
            .commit_with_package()
            .map_err(Error::from)?;
        let package = self.package;
        commit_from_body(package, body_commit)
    }

    /// Publish the checked package commit while reporting deterministic,
    /// content-free phase boundaries to `observer` (the opt-in profiling
    /// equivalent of [`Self::commit`]).
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::commit`].
    #[cfg(feature = "performance-diagnostics")]
    pub fn commit_profiled(
        self,
        observer: impl FnMut(body_text::DiagnosticEvent),
    ) -> Result<Commit> {
        self.commit_profiled_with_cfb_observer(observer, |_| {})
    }

    /// [`Self::commit_profiled`] with a separate in-memory CFB observer.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::commit`].
    #[cfg(feature = "performance-diagnostics")]
    pub fn commit_profiled_with_cfb_observer(
        mut self,
        observer: impl FnMut(body_text::DiagnosticEvent),
        cfb_observer: impl FnMut(body_text::CfbParseEvent),
    ) -> Result<Commit> {
        let (snapshot, body_patch) = self
            .take_inner()?
            .commit_profiled_with_cfb_observer(observer, cfb_observer)
            .map_err(Error::from)?
            .into_parts();
        let source = self.package;
        let patch = Patch::from_body(body_patch);
        let diagnostics = Diagnostics::from_patch(&patch.inner);
        let profile = EditProfile::from_snapshot(&snapshot);
        let package = if snapshot.bytes() == source.to_bytes().map_err(Error::Package)?.as_slice() {
            source
        } else {
            Package::from_bytes_with_options(
                snapshot.bytes(),
                OpenOptions::default()
                    .source_limits(snapshot.source_limits())
                    .leniency(source.leniency()),
            )
            .map_err(Error::Package)?
        };
        Ok(Commit::new(
            package.with_edit_profile(profile),
            patch,
            diagnostics,
        ))
    }
}

impl Package {
    /// Materializes the body-text snapshot of this package under `options`.
    /// The snapshot is memoized after the first body operation: an immutable
    /// snapshot of the same bytes under the same options is the value a fresh
    /// open returns, and every edit starts from its own transaction over it. A
    /// memo hit still checks that the retained source has its captured
    /// version, as the copy does.
    pub(crate) fn body_snapshot(&self, options: &EditOptions) -> Result<body_text::Snapshot> {
        self.body_source(
            body_text::OpenOptions::default()
                .source_limits(self.source_limits())
                .revision_limits(options.revision_limits)
                .transaction_limits(options.transaction_limits)
                .protection_policy(options.protection_policy.clone()),
        )
    }

    fn body_source(&self, options: body_text::OpenOptions) -> Result<body_text::Snapshot> {
        let memo = || {
            self.body_source
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
        };
        let held = memo()
            .as_ref()
            .filter(|(held, _)| *held == options)
            .map(|(_, snapshot)| snapshot.clone());
        if let Some(snapshot) = held {
            self.check_source_current().map_err(Error::Package)?;
            return Ok(snapshot);
        }
        let bytes = self.to_bytes().map_err(Error::Package)?;
        // `options` carry this package's source limits (`body_snapshot`).
        let snapshot =
            body_text::Snapshot::from_package_bytes(bytes, options.clone(), self.strict_document)
                .map_err(Error::from)?;
        *memo() = Some((options, snapshot.clone()));
        Ok(snapshot)
    }

    fn with_edit_profile(mut self, profile: EditProfile) -> Self {
        self.edit_profile = Some(profile);
        self
    }

    /// Start the canonical primary DOC package transaction.
    ///
    /// The package remains immutable. Construction retains only the package
    /// and options; the first body operation copies and revalidates its exact
    /// bytes through the existing body-text owner. Dropping the returned edit
    /// discards those deferred or staged changes.
    pub fn edit(&self) -> Result<Edit<'static>> {
        Edit::from_package_with_options(self.clone(), self.retained_edit_options())
    }

    /// The edit options this package retains from the edit that produced it
    /// (the defaults for a freshly opened package).
    pub(crate) fn retained_edit_options(&self) -> EditOptions {
        self.edit_profile
            .as_ref()
            .map(EditProfile::options)
            .unwrap_or_default()
    }

    /// This package retaining `options` as its edit profile, as a scoped
    /// family commit publishes it.
    pub(crate) fn with_edit_options(self, options: &EditOptions) -> Self {
        self.with_edit_profile(EditProfile {
            revision_limits: options.revision_limits,
            transaction_limits: options.transaction_limits,
            protection_policy: options.protection_policy.clone(),
        })
    }

    /// `options` constrained by the edit profile this package retains.
    pub(crate) fn constrained_edit_options(&self, options: EditOptions) -> EditOptions {
        match self.edit_profile.as_ref() {
            Some(profile) => profile.constrain_options(options),
            None => options,
        }
    }

    /// Start a primary transaction under one explicit edit profile.
    pub fn edit_with_options(&self, options: EditOptions) -> Result<Edit<'static>> {
        Edit::from_package_with_options(self.clone(), self.constrained_edit_options(options))
    }

    /// Apply a canonical exact-source patch without mutating this package.
    pub fn apply(&self, patch: &Patch) -> Result<Commit> {
        let bytes = self.to_bytes().map_err(Error::Package)?;
        let profile = self.edit_profile.clone().unwrap_or_else(|| {
            let mut profile = EditProfile::from_patch(&patch.inner);
            // A pristine package has no caller-granted edit capability.
            // Reuse the patch's numeric bounds so a wide explicit edit
            // does not fall back to the body default, while retaining the
            // destination's strict protection authority.
            profile.protection_policy = ProtectionPolicy::default();
            profile
        });
        let source = body_text::Snapshot::from_bytes_with_options(
            bytes,
            body_text::OpenOptions::default()
                .source_limits(self.source_limits())
                .revision_limits(profile.revision_limits)
                .transaction_limits(profile.transaction_limits)
                .protection_policy(profile.protection_policy),
        )
        .map_err(Error::from)?;
        let next = patch.inner.apply(&source).map_err(Error::from)?;
        let next_profile = EditProfile::from_snapshot(&next);
        let diagnostics = Diagnostics::from_patch(&patch.inner);
        if next.bytes() == source.bytes() {
            return Ok(Commit::new(
                self.clone().with_edit_profile(next_profile),
                patch.clone(),
                diagnostics,
            ));
        }
        let package = Package::from_bytes_with_options(
            next.bytes(),
            OpenOptions::default()
                .source_limits(next.source_limits())
                .leniency(self.leniency()),
        )
        .map_err(Error::Package)?
        .with_edit_profile(next_profile);
        Ok(Commit::new(package, patch.clone(), diagnostics))
    }
}

#[cfg(test)]
mod tests {
    #[cfg(feature = "performance-diagnostics")]
    use super::Edit;
    use super::{Diagnostics, EditOptions, Error, LimitKind};
    use crate::Builder;
    use crate::body_text::{self, TextTarget, TransactionLimits};
    use crate::parts::fib::FileInformationBlock;
    use crate::parts::protection::{ProtectionAuthorization, ProtectionPolicy};
    use crate::tracked_revision::Limits as RevisionLimits;
    use litchi_core::Position;
    use litchi_ole_common::object::{
        OpenOptions as ObjectOpenOptions, Package as ObjectPackage, Targets,
    };
    use std::io::Cursor;

    #[cfg(feature = "performance-diagnostics")]
    #[test]
    fn ordinary_edit_keeps_send_and_sync_auto_traits() {
        fn assert_send_sync<T: Send + Sync>() {}

        assert_send_sync::<Edit<'static>>();
    }

    #[cfg(feature = "performance-diagnostics")]
    #[allow(dead_code)]
    fn accepts_borrowed_non_send_profiled_observer<'a>(
        package: &'a crate::Package,
        marker: &'a std::rc::Rc<std::cell::Cell<usize>>,
    ) -> super::Result<Edit<'a, super::ProfiledOpen<'a>>> {
        package.edit_profiled(
            EditOptions::default(),
            move |_| marker.set(marker.get().saturating_add(1)),
            |_| {},
        )
    }

    fn package() -> crate::Package {
        package_with_paragraphs(1)
    }

    fn package_with_paragraphs(count: usize) -> crate::Package {
        let mut builder = Builder::new();
        for _ in 0..count {
            builder.add_paragraph("before").expect("fixture paragraph");
        }
        builder.build().expect("fixture package")
    }

    fn protected_package_bytes() -> Vec<u8> {
        let mut builder = Builder::new();
        builder.add_paragraph("before").expect("fixture paragraph");
        let mut output = Cursor::new(Vec::new());
        builder
            .build()
            .expect("fixture package")
            .write_to(&mut output)
            .expect("fixture DOC");
        let bytes = output.into_inner();
        let mut package = ObjectPackage::open_with_options(
            bytes,
            ObjectOpenOptions::new()
                .targets(Targets::default())
                .limits(RevisionLimits::default()),
        )
        .expect("fixture package should open")
        .edit()
        .expect("fixture package should be editable");
        let word_path = ["WordDocument".to_owned()];
        let word = package.stream(&word_path).expect("WordDocument");
        let fib = FileInformationBlock::parse(word).expect("fixture FIB");
        let table_name = if fib.which_table_stream() {
            "1Table"
        } else {
            "0Table"
        };
        let table_path = [table_name.to_owned()];
        let mut word = word.to_vec();
        let mut table = package.stream(&table_path).expect("table stream").to_vec();
        let (offset, length) = fib.get_table_pointer(31).expect("DOP pointer");
        let offset = usize::try_from(offset).expect("DOP offset");
        let length = usize::try_from(length).expect("DOP length");
        let mut dop = crate::parts::document_properties::DocumentProperties::word97_writer_bytes(
            false, false, false,
        );
        dop.resize(594, 0);
        if length < dop.len() {
            let insertion = offset + length;
            let extra = dop.len() - length;
            table.splice(insertion..insertion, std::iter::repeat_n(0, extra));
            let count = fib.table_pointer_count().expect("FIB pointer count");
            for index in 0..count {
                let pointer = 154 + index * 8;
                let current = usize::try_from(u32::from_le_bytes(
                    word[pointer..pointer + 4]
                        .try_into()
                        .expect("FIB pointer offset"),
                ))
                .expect("FIB pointer offset");
                if current >= insertion {
                    let shifted = u32::try_from(current + extra).expect("shifted FIB pointer");
                    word[pointer..pointer + 4].copy_from_slice(&shifted.to_le_bytes());
                }
            }
        }
        assert!(table.len() >= offset + dop.len());
        table[offset..offset + dop.len()].copy_from_slice(&dop);
        let pointer = 154 + 31 * 8;
        word[pointer + 4..pointer + 8]
            .copy_from_slice(&u32::try_from(dop.len()).expect("DOP length").to_le_bytes());
        package
            .put_stream(&word_path, word)
            .expect("WordDocument stream");
        package
            .put_stream(&table_path, table)
            .expect("table stream");
        let mut table = package.stream(&table_path).expect("table stream").to_vec();
        let word = package.stream(&word_path).expect("WordDocument");
        let fib = FileInformationBlock::parse(word).expect("updated FIB");
        let (offset, _) = fib.get_table_pointer(31).expect("updated DOP pointer");
        let offset = usize::try_from(offset).expect("updated DOP offset");
        table[offset + 6] |= 0x10;
        package
            .put_stream(&table_path, table)
            .expect("protected table stream");
        crate::package::object_edit_bytes(&package).expect("protected fixture publication")
    }

    #[test]
    fn primary_commit_exposes_canonical_triad_and_inverse() {
        let source = package();
        let source_bytes = source.to_bytes().expect("source bytes");
        let mut edit = source.edit().expect("edit");
        edit.replace_paragraph(Position::new(0), "after")
            .expect("replace");
        let commit = edit.commit().expect("commit");
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operation_count(), 1);
        assert_ne!(commit.package().to_bytes().unwrap(), source_bytes);
        let restored = commit
            .package()
            .apply(&commit.patch().inverse())
            .expect("inverse");
        assert_eq!(restored.package().to_bytes().unwrap(), source_bytes);
    }

    #[test]
    fn primary_noop_preserves_bytes_and_is_reported() {
        let source = package();
        let mut edit = source.edit().expect("edit");
        edit.replace_text(TextTarget::body_paragraph(Position::new(0)), "before")
            .expect("no-op");
        let commit = edit.commit().expect("commit");
        assert_eq!(
            *commit.diagnostics(),
            Diagnostics {
                changed: false,
                operation_count: 0,
                auxiliary_changed: false,
            }
        );
        assert_eq!(
            commit.package().to_bytes().unwrap(),
            source.to_bytes().unwrap()
        );
    }

    #[test]
    fn stale_patch_is_rejected_without_mutating_source() {
        let source = package();
        let mut first = source.edit().expect("edit");
        first
            .replace_paragraph(Position::new(0), "first")
            .expect("replace");
        let patch = first.commit().expect("commit").patch().clone();

        let mut second = source.edit().expect("edit");
        second
            .replace_paragraph(Position::new(0), "second")
            .expect("replace");
        let changed = second.commit().expect("commit").into_package();
        let before = changed.to_bytes().expect("bytes");
        let error = changed.apply(&patch).expect_err("stale patch");
        assert!(matches!(
            error,
            Error::Body(body_text::Error::PatchConflict)
        ));
        assert_eq!(changed.to_bytes().unwrap(), before);
    }

    /// Record 0311: edits of one package share the body snapshot opened for
    /// their options, never one opened under other options, and publish what
    /// an edit from a fresh open publishes.
    #[test]
    fn edits_share_the_body_snapshot_opened_under_their_options() {
        let source = package_with_paragraphs(3);
        let commit = |package: &crate::Package, options: EditOptions| {
            let mut edit = package.edit_with_options(options).expect("edit");
            edit.replace_paragraph(Position::new(1), "after")
                .expect("replacement");
            edit.commit()
                .expect("commit")
                .into_package()
                .to_bytes()
                .unwrap()
        };
        let first = commit(&source, EditOptions::default());
        let held = |package: &crate::Package| {
            package
                .body_source
                .lock()
                .unwrap()
                .as_ref()
                .map(|(_, snapshot)| snapshot.bytes_shared())
        };
        let shared = held(&source).expect("the first edit keeps its snapshot");
        assert_eq!(&*shared, source.to_bytes().unwrap().as_slice());
        // A second edit, from a clone too, reuses that snapshot.
        assert_eq!(commit(&source.clone(), EditOptions::default()), first);
        assert!(std::sync::Arc::ptr_eq(&shared, &held(&source).unwrap()));
        // Other options open their own snapshot.
        let wide = EditOptions::default().transaction_limits(TransactionLimits::new(
            300,
            1024 * 1024,
            4 * 1024 * 1024,
        ));
        assert_eq!(commit(&source, wide), first);
        assert!(!std::sync::Arc::ptr_eq(&shared, &held(&source).unwrap()));
        // A fresh open publishes the same bytes.
        let fresh = crate::Package::from_bytes(&source.to_bytes().unwrap()).unwrap();
        assert_eq!(commit(&fresh, EditOptions::default()), first);
    }

    /// Record 0377a: a package whose document is the strict reader's parse of
    /// its bytes opens its body snapshot without parsing them again, and gets
    /// the snapshot a full open of those bytes returns. A lenient open keeps
    /// the full reader gate; a commit's document is the commit's strict parse.
    #[test]
    fn a_strictly_parsed_package_opens_its_body_snapshot_without_a_second_parse() {
        let source = package_with_paragraphs(3);
        let bytes = source.to_bytes().unwrap();
        assert!(source.strict_document);
        assert!(crate::Package::from_bytes(&bytes).unwrap().strict_document);

        let options = EditOptions::default();
        let opened = source.body_snapshot(&options).expect("package snapshot");
        let full = body_text::Snapshot::from_bytes_with_options(
            bytes.clone(),
            body_text::OpenOptions::default()
                .source_limits(source.source_limits())
                .revision_limits(options.revision_limits)
                .transaction_limits(options.transaction_limits)
                .protection_policy(options.protection_policy.clone()),
        )
        .expect("full snapshot");
        assert_eq!(opened.bytes(), full.bytes());
        assert_eq!(opened.fingerprint(), full.fingerprint());
        assert_eq!(opened.source_limits(), full.source_limits());

        let lenient = crate::Package::from_bytes_with_options(
            &bytes,
            super::OpenOptions::default().leniency(crate::Leniency::TolerateStylesheetDefects),
        )
        .unwrap();
        assert!(!lenient.strict_document);
        let mut edit = lenient.edit().expect("edit");
        edit.replace_paragraph(Position::new(1), "after")
            .expect("replacement");
        let committed = edit.commit().expect("commit").into_package();
        assert_eq!(
            committed.leniency(),
            crate::Leniency::TolerateStylesheetDefects
        );
        assert!(committed.strict_document);
    }

    #[test]
    fn retained_profile_allows_more_than_default_operations_through_inverse() {
        let source = package_with_paragraphs(257);
        let transaction_limits = TransactionLimits::new(300, 1024 * 1024, 4 * 1024 * 1024);
        let mut edit = source
            .edit_with_options(EditOptions::default().transaction_limits(transaction_limits))
            .expect("wide edit");
        for position in 0..257 {
            edit.replace_paragraph(Position::new(position), "after")
                .expect("wide replacement");
        }
        let commit = edit.commit().expect("wide commit");
        let patch = commit.patch().clone();
        let changed = commit.into_package();
        let restored = changed
            .apply(&patch.inverse())
            .expect("inverse with retained profile");
        assert_eq!(
            restored.package().to_bytes().unwrap(),
            source.to_bytes().unwrap()
        );

        let mut follow_up = restored.package().edit().expect("retained profile edit");
        for position in 0..257 {
            follow_up
                .replace_paragraph(Position::new(position), "again")
                .expect("retained profile replacement");
        }
    }

    #[test]
    fn apply_retains_the_narrower_destination_transaction_profile() {
        let source = package_with_paragraphs(3);
        let wide_limits = TransactionLimits::new(3, 1024 * 1024, 4 * 1024 * 1024);
        let mut wide_edit = source
            .edit_with_options(EditOptions::default().transaction_limits(wide_limits))
            .expect("wide edit");
        for position in 0..3 {
            wide_edit
                .replace_paragraph(Position::new(position), "wide")
                .expect("wide replacement");
        }
        let wide_patch = wide_edit.commit().expect("wide commit").patch().clone();

        let narrow_limits = TransactionLimits::new(2, 1024 * 1024, 4 * 1024 * 1024)
            .with_max_retained_render_bytes(1);
        let narrow = source
            .edit_with_options(EditOptions::default().transaction_limits(narrow_limits))
            .expect("narrow edit")
            .commit()
            .expect("narrow no-op commit")
            .into_package();
        let applied = narrow
            .apply(&wide_patch)
            .expect("wide patch on narrow source");
        let wider_options = EditOptions::default().transaction_limits(
            TransactionLimits::new(99, 4 * 1024 * 1024, 16 * 1024 * 1024)
                .with_max_retained_render_bytes(4 * 1024 * 1024),
        );
        let mut follow_up = applied
            .package()
            .edit_with_options(wider_options)
            .expect("narrow retained profile edit");
        follow_up
            .replace_paragraph(Position::new(0), "narrow")
            .expect("first narrow replacement");
        follow_up
            .replace_paragraph(Position::new(1), "narrow")
            .expect("second narrow replacement");
        let error = follow_up
            .replace_paragraph(Position::new(2), "narrow")
            .expect_err("narrow operation limit");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::Operations,
                ..
            }
        ));
        if let Some(retained) = follow_up.retained_render_bytes() {
            assert!(retained <= 1);
        }
    }

    #[test]
    fn equal_reopened_package_accepts_exact_patch_source() {
        let source = package();
        let mut edit = source.edit().expect("edit");
        edit.replace_paragraph(Position::new(0), "after")
            .expect("replace");
        let patch = edit.commit().expect("commit").patch().clone();
        let reopened = crate::Package::from_bytes(&source.to_bytes().unwrap()).unwrap();
        let applied = reopened.apply(&patch).expect("equal fresh source");
        assert_eq!(
            applied.package().to_bytes().unwrap(),
            patch.after_bytes(),
            "exact bytes authorize a separately reopened source"
        );
    }

    #[test]
    fn fresh_reopened_protected_package_keeps_destination_protection_authority() {
        let source_bytes = protected_package_bytes();
        let source = crate::Package::from_bytes(&source_bytes).expect("protected source");
        let authorization = ProtectionAuthorization::audited(
            "package-edit-tests",
            "exercise an explicitly authorized protected primary edit",
        )
        .expect("authorization");
        let policy = ProtectionPolicy::allow_protected(authorization);
        let mut edit = source
            .edit_with_options(EditOptions::default().protection_policy(policy))
            .expect("authorized edit");
        edit.replace_paragraph(Position::new(0), "after")
            .expect("protected replacement");
        let commit = edit.commit().expect("protected commit");
        let patch = commit.patch().clone();
        let restored = commit
            .package()
            .apply(&patch.inverse())
            .expect("authorized package can apply its inverse");
        assert_eq!(restored.package().to_bytes().unwrap(), source_bytes);

        let reopened = crate::Package::from_bytes(&source_bytes).expect("fresh source");
        let error = reopened
            .apply(&patch)
            .expect_err("fresh destination must reject protected publication");
        assert!(matches!(
            error,
            Error::Body(body_text::Error::Refused(body_text::Refusal::Protected))
        ));
    }

    #[test]
    fn commit_rejects_candidate_over_outer_package_limit_before_public_parse() {
        let source_bytes = package().to_bytes().expect("source bytes");
        let limits = crate::SourceLimits::builder()
            .max_input_bytes(source_bytes.len())
            .expect("source-sized package limit")
            .build()
            .expect("source-sized package limits");
        let source = crate::Package::from_bytes_with_options(
            &source_bytes,
            crate::OpenOptions::default().source_limits(limits),
        )
        .expect("limited source");
        let mut edit = source.edit().expect("edit");
        edit.replace_paragraph(Position::new(0), &"x".repeat(100_000))
            .expect("large replacement");
        let error = edit
            .commit()
            .expect_err("candidate must exceed the outer package limit");
        assert!(matches!(
            error,
            Error::Body(body_text::Error::Invalid(
                crate::package::Error::Limit { .. }
            ))
        ));
    }
}
