//! The external-hyperlink detachment family of a source-backed package
//! (record 0355a): unwrap the main document's external hyperlink
//! wrappers, keeping their visible children and the relationship records.
//!
//! [`Package::edit_external_hyperlink_detachment`](crate::source_backed::Package::edit_external_hyperlink_detachment)
//! stages an [`Edit`] whose commit returns a new source-backed package;
//! writing it rewrites only the main-document part.
//! [`Package::apply_external_hyperlink_detachment`](crate::source_backed::Package::apply_external_hyperlink_detachment)
//! replays a [`Patch`].

use crate::Result;
use crate::sanitize::{self as record, EffectReport};
use crate::source_backed::Package;
use crate::source_backed::deferred::{Change, source_family};

/// The detached main-document patch the family publishes.
#[derive(Clone)]
pub(crate) struct DetachmentChange(record::Patch);

impl Change for DetachmentChange {
    type Error = crate::Error;

    fn is_noop(&self) -> bool {
        self.0.is_noop()
    }

    fn inverse(&self) -> Self {
        Self(self.0.inverse())
    }

    fn publish(&self, package: &Package) -> Result<Package> {
        let (main, target) = package.external_hyperlink_detachment_target(&self.0)?;
        if self.is_noop() {
            return Ok(package.clone());
        }
        package.deferred_part_overlay(main, target)
    }
}

source_family!("external-hyperlink detachment", DetachmentChange);

impl From<record::Patch> for Patch {
    /// Replay a detached record patch onto a package whose main document
    /// and relationships hold its before state.
    fn from(patch: record::Patch) -> Self {
        Self {
            change: DetachmentChange(patch),
            restore: None,
        }
    }
}

/// The budgets of this family, reported by
/// [`crate::Error::ExternalHyperlinkDetachmentLimit`].
pub use crate::sanitize::LimitKind;

/// A staged detachment of every external main-document hyperlink wrapper.
/// The source is never changed: [`Self::commit`] returns a new package.
#[derive(Clone)]
pub struct Edit {
    package: Package,
    plan: record::SanitizePlan,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
    /// The package the edit is staged over.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// The predicted effects, without changing anything.
    #[must_use]
    pub const fn effect_report(&self) -> EffectReport {
        self.plan.effect_report()
    }

    /// Commit the detachment as a new source-backed package.
    ///
    /// # Errors
    ///
    /// Returns an error when the rewrite would change visible text or keep a
    /// selected wrapper, and the source, protection or signature errors.
    pub fn commit(self) -> Result<Commit> {
        let patch = self.plan.apply()?;
        commit_change(&self.package, DetachmentChange(patch))
    }
}

impl Package {
    /// Stage the detachment of every external main-document hyperlink
    /// wrapper.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::external_hyperlink_sanitization_snapshot`].
    pub fn edit_external_hyperlink_detachment(&self) -> Result<Edit> {
        Ok(Edit {
            package: self.clone(),
            plan: self.plan_external_hyperlink_detachment()?,
        })
    }

    /// Replay a detachment patch onto this package, which must hold the
    /// patch's before state.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] for another main-document or
    /// relationship state.
    pub fn apply_external_hyperlink_detachment(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}
