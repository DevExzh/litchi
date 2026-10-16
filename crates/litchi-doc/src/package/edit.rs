//! Canonical primary DOC package transaction.
//!
//! [`Package`] is the immutable owner used by this module.  The transaction
//! delegates its checked semantic operations to [`crate::body_text`], which
//! remains the source-preserving owner for ordinary DOC text and its bounded
//! dependency closures.  The package boundary adds the canonical
//! package-owned publication shape without introducing a second edit engine.
//!
//! The current bridge copies and revalidates the retained source when an edit
//! is opened. Changed candidates transfer the body engine's already validated
//! CFB and document projection into the package owner; exact no-ops retain the
//! original package. The body engine's exact-source, protection, and
//! full-candidate validation gates remain authoritative at every publication
//! boundary.

use super::{Error as PackageError, OpenOptions, Package, model::EditProfile};
use crate::body_text;
use crate::parts::dofr::DofrPatch;
use crate::parts::protection::ProtectionPolicy;
use crate::parts::saved_selection::{SavedSelectionPatch, SavedSelectionTransaction};
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
    pub const fn with_revision_limits(mut self, limits: Limits) -> Self {
        self.revision_limits = limits;
        self
    }

    /// Select finite staged-operation and replacement limits.
    #[must_use]
    pub const fn with_transaction_limits(mut self, limits: body_text::TransactionLimits) -> Self {
        self.transaction_limits = limits;
        self
    }

    /// Select the explicit policy for changed protected-source publication.
    #[must_use]
    pub fn with_protection_policy(mut self, policy: ProtectionPolicy) -> Self {
        self.protection_policy = policy;
        self
    }

    /// Return the tracked-revision limits.
    #[must_use]
    pub const fn revision_limits(&self) -> Limits {
        self.revision_limits
    }

    /// Return the transaction limits.
    #[must_use]
    pub const fn transaction_limits(&self) -> body_text::TransactionLimits {
        self.transaction_limits
    }

    /// Return the protected-publication policy.
    #[must_use]
    pub fn protection_policy(&self) -> ProtectionPolicy {
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
    inner: body_text::Patch,
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
    pub const fn transaction_limits(&self) -> body_text::TransactionLimits {
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

/// A package-owned commit carrying the canonical publication triad.
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
    fn from_body(source: Package, body_commit: body_text::PackageCommit) -> Result<Self> {
        let body_text::PackageCommit {
            snapshot,
            patch: body_patch,
            validated,
        } = body_commit;
        let patch = Patch::from_body(body_patch);
        let diagnostics = Diagnostics::from_patch(&patch.inner);
        let package = match validated {
            None => source,
            Some(validated) => Package::from_validated_ole(
                validated.ole,
                validated.document,
                snapshot.source_limits(),
                source.leniency(),
            )
            .map_err(Error::Package)?,
        };
        Ok(Self {
            package: package.with_edit_profile(EditProfile::from_snapshot(&snapshot)),
            patch,
            diagnostics,
        })
    }

    /// Published immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Content-free publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit into its package and reversible patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

/// Primary immutable DOC package transaction.
pub struct Edit {
    package: Package,
    inner: body_text::Edit,
}

impl fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("package", &self.package)
            .field("retained_render_bytes", &self.retained_render_bytes())
            .finish_non_exhaustive()
    }
}

impl Edit {
    pub(crate) fn from_package_with_options(
        package: Package,
        options: EditOptions,
    ) -> Result<Self> {
        if package.document().fib().is_encrypted() {
            return Err(Error::Body(body_text::Error::Refused(
                body_text::Refusal::Encrypted,
            )));
        }
        let source = package.body_source(
            body_text::OpenOptions::default()
                .source_limits(package.source_limits())
                .with_revision_limits(options.revision_limits)
                .with_transaction_limits(options.transaction_limits)
                .with_protection_policy(options.protection_policy),
        )?;
        let inner = source.edit().map_err(Error::from)?;
        Ok(Self { package, inner })
    }

    /// Immutable package snapshot that authorizes this transaction.
    #[must_use]
    pub const fn source(&self) -> &Package {
        &self.package
    }

    /// Exact source bytes retained by this staged edit.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        self.inner.source_bytes()
    }

    /// Borrowed semantic changes staged so far, including authored text.
    #[must_use]
    pub fn changes(&self) -> impl ExactSizeIterator<Item = body_text::ChangeRef<'_>> {
        self.inner.changes()
    }

    /// Allocation capacity held by the validated candidate, if retained.
    #[must_use]
    pub fn retained_render_bytes(&self) -> Option<usize> {
        self.inner.retained_render_bytes()
    }

    /// Release the retained candidate without changing staged semantics.
    pub fn release_retained_render(&mut self) {
        self.inner.release_retained_render();
    }

    /// Apply a checked `Selsf` patch to this transaction.
    pub fn apply_saved_selection_patch(&mut self, patch: &SavedSelectionPatch) -> Result<bool> {
        self.inner
            .apply_saved_selection_patch(patch)
            .map_err(Error::from)
    }

    /// Edit the optional `Selsf` record through its typed transaction.
    pub fn edit_saved_selection(
        &mut self,
        edit: impl FnOnce(&mut SavedSelectionTransaction) -> crate::package::Result<()>,
    ) -> Result<bool> {
        self.inner.edit_saved_selection(edit).map_err(Error::from)
    }

    /// Apply a checked `RgDofr` patch to this transaction.
    pub fn apply_dofr_patch(&mut self, patch: &DofrPatch) -> Result<bool> {
        self.inner.apply_dofr_patch(patch).map_err(Error::from)
    }

    /// Replace one complete checked `RgDofr` record.
    pub fn replace_dofr_record(&mut self, index: usize, replacement: &[u8]) -> Result<bool> {
        self.inner
            .replace_dofr_record(index, replacement)
            .map_err(Error::from)
    }

    /// Replace one ordinary main-story paragraph.
    pub fn replace_paragraph(&mut self, position: Position, replacement: &str) -> Result<()> {
        self.inner
            .replace_paragraph(position, replacement)
            .map_err(Error::from)
    }

    /// Replace one checked body, table-cell, or field target.
    pub fn replace_text(&mut self, target: body_text::TextTarget, replacement: &str) -> Result<()> {
        self.inner
            .replace_text(target, replacement)
            .map_err(Error::from)
    }

    /// Set direct bold on one ordinary paragraph.
    pub fn set_paragraph_bold(&mut self, position: Position, enabled: bool) -> Result<()> {
        self.inner
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
        self.inner
            .set_character_property(target, property, enabled)
            .map_err(Error::from)
    }

    /// Set one valued direct character property (size, color or font).
    pub fn set_character_value(
        &mut self,
        target: body_text::TextTarget,
        value: body_text::CharacterValue<'_>,
    ) -> Result<()> {
        self.inner
            .set_character_value(target, value)
            .map_err(Error::from)
    }

    /// Insert one ordinary main-story paragraph before `position`, or after
    /// the last one when `position` is the paragraph count.
    pub fn insert_paragraph(&mut self, position: Position, text: &str) -> Result<()> {
        self.inner
            .insert_paragraph(position, text)
            .map_err(Error::from)
    }

    /// Remove one ordinary main-story paragraph.
    pub fn remove_paragraph(&mut self, position: Position) -> Result<()> {
        self.inner.remove_paragraph(position).map_err(Error::from)
    }

    /// Move one ordinary main-story paragraph so that it ends at `to`,
    /// counted after it is detached.
    pub fn move_paragraph(&mut self, from: Position, to: Position) -> Result<()> {
        self.inner.move_paragraph(from, to).map_err(Error::from)
    }

    /// Apply a dependency-checked text transfer.
    pub fn apply_transfer(&mut self, plan: &body_text::TransferPlan) -> Result<()> {
        self.inner.apply_transfer(plan).map_err(Error::from)
    }

    /// Apply a dependency-closed embedded-object transfer.
    pub fn apply_embedded_transfer(
        &mut self,
        plan: &body_text::EmbeddedTransferPlan,
    ) -> Result<()> {
        self.inner
            .apply_embedded_transfer(plan)
            .map_err(Error::from)
    }

    /// Apply a checked inline or floating picture transfer.
    pub fn apply_picture_transfer(&mut self, plan: &body_text::PictureTransferPlan) -> Result<()> {
        self.inner.apply_picture_transfer(plan).map_err(Error::from)
    }

    /// Add one inert embedded object through its dependency owner.
    pub fn add_embedded_object(
        &mut self,
        options: crate::embedded_object::WriteOptions,
    ) -> Result<()> {
        self.inner.add_embedded_object(options).map_err(Error::from)
    }

    /// Remove one managed embedded object and its dependency closure.
    pub fn remove_embedded_object(&mut self, storage_id: u32) -> Result<()> {
        self.inner
            .remove_embedded_object(storage_id)
            .map_err(Error::from)
    }

    /// Accept or reject one safe revision mark.
    pub fn dispose_revision(
        &mut self,
        position: Position,
        disposition: body_text::RevisionDisposition,
    ) -> Result<()> {
        self.inner
            .dispose_revision(position, disposition)
            .map_err(Error::from)
    }

    /// Change the passive display-as-icon bit of one managed object.
    pub fn set_embedded_display_as_icon(&mut self, storage_id: u32, enabled: bool) -> Result<()> {
        self.inner
            .set_embedded_display_as_icon(storage_id, enabled)
            .map_err(Error::from)
    }

    /// Publish the checked package commit and canonical triad.
    pub fn commit(self) -> Result<Commit> {
        let package = self.package;
        let body_commit = self.inner.commit_with_package().map_err(Error::from)?;
        Commit::from_body(package, body_commit)
    }
}

impl Package {
    /// The body-text snapshot of this package's exact bytes under `options`.
    ///
    /// Opening one copies and parses the whole file, so the last one opened
    /// is kept (record 0311): an immutable snapshot of the same bytes under
    /// the same options is the value a fresh open returns, and every edit
    /// starts from its own transaction over it. A memo hit still checks that
    /// the retained source has its captured version, as the copy does.
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
        let snapshot = body_text::Snapshot::from_bytes_with_options(bytes, options.clone())
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
    /// The package remains immutable.  Staging copies and revalidates its
    /// exact bytes through the existing body-text owner; dropping the returned
    /// edit is the rollback operation.
    pub fn edit(&self) -> Result<Edit> {
        let options = self
            .edit_profile
            .as_ref()
            .map(EditProfile::options)
            .unwrap_or_default();
        Edit::from_package_with_options(self.clone(), options)
    }

    /// Start a primary transaction under one explicit edit profile.
    pub fn edit_with_options(&self, options: EditOptions) -> Result<Edit> {
        let options = match self.edit_profile.as_ref() {
            Some(profile) => profile.constrain_options(options),
            None => options,
        };
        Edit::from_package_with_options(self.clone(), options)
    }

    /// Apply a canonical exact-source patch without mutating this package.
    pub fn apply(&self, patch: &Patch) -> Result<Self> {
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
                .with_revision_limits(profile.revision_limits)
                .with_transaction_limits(profile.transaction_limits)
                .with_protection_policy(profile.protection_policy),
        )
        .map_err(Error::from)?;
        let next = patch.inner.apply(&source).map_err(Error::from)?;
        let next_profile = EditProfile::from_snapshot(&next);
        if next.bytes() == source.bytes() {
            return Ok(self.clone().with_edit_profile(next_profile));
        }
        Package::from_bytes_with_options(
            next.bytes(),
            OpenOptions::default()
                .source_limits(next.source_limits())
                .leniency(self.leniency()),
        )
        .map_err(Error::Package)
        .map(|package| package.with_edit_profile(next_profile))
    }
}

#[cfg(test)]
mod tests {
    use super::{Diagnostics, EditOptions, Error, LimitKind};
    use crate::Builder;
    use crate::body_text::{self, TextTarget, TransactionLimits};
    use crate::parts::fib::FileInformationBlock;
    use crate::parts::protection::{ProtectionAuthorization, ProtectionPolicy};
    use crate::tracked_revision::Limits as RevisionLimits;
    use litchi_core::Position;
    use litchi_ole_common::object::{Editor as PackageEditor, Targets};
    use std::io::Cursor;

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
        let mut package = PackageEditor::open(bytes, Targets::default(), RevisionLimits::default())
            .expect("fixture package should open");
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
        package.finish().expect("protected fixture publication")
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
        assert_eq!(restored.to_bytes().unwrap(), source_bytes);
    }

    #[test]
    fn primary_noop_preserves_bytes_and_is_reported() {
        let source = package();
        let mut edit = source.edit().expect("edit");
        edit.replace_text(TextTarget::body_paragraph(Position::new(0)), "before")
            .expect("no-op");
        let commit = edit.commit().expect("commit");
        assert_eq!(
            commit.diagnostics(),
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
        assert!(matches!(error, Error::Body(body_text::Error::Conflict)));
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
        let wide = EditOptions::default().with_transaction_limits(TransactionLimits::new(
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

    #[test]
    fn retained_profile_allows_more_than_default_operations_through_inverse() {
        let source = package_with_paragraphs(257);
        let transaction_limits = TransactionLimits::new(300, 1024 * 1024, 4 * 1024 * 1024);
        let mut edit = source
            .edit_with_options(EditOptions::default().with_transaction_limits(transaction_limits))
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
        assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());

        let mut follow_up = restored.edit().expect("retained profile edit");
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
            .edit_with_options(EditOptions::default().with_transaction_limits(wide_limits))
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
            .edit_with_options(EditOptions::default().with_transaction_limits(narrow_limits))
            .expect("narrow edit")
            .commit()
            .expect("narrow no-op commit")
            .into_package();
        let applied = narrow
            .apply(&wide_patch)
            .expect("wide patch on narrow source");
        let wider_options = EditOptions::default().with_transaction_limits(
            TransactionLimits::new(99, 4 * 1024 * 1024, 16 * 1024 * 1024)
                .with_max_retained_render_bytes(4 * 1024 * 1024),
        );
        let mut follow_up = applied
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
            applied.to_bytes().unwrap(),
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
            .edit_with_options(EditOptions::default().with_protection_policy(policy))
            .expect("authorized edit");
        edit.replace_paragraph(Position::new(0), "after")
            .expect("protected replacement");
        let commit = edit.commit().expect("protected commit");
        let patch = commit.patch().clone();
        let restored = commit
            .package()
            .apply(&patch.inverse())
            .expect("authorized package can apply its inverse");
        assert_eq!(restored.to_bytes().unwrap(), source_bytes);

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
