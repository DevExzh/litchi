//! Reversible source-checked Custom XML package patches.

use crate::Result;
use crate::signature::refuse_signed_change;
use litchi_opc::OpcPackage;

use super::package;
use super::snapshot::Snapshot;

/// A complete, reversible replacement of one Custom XML package graph.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source graph required before this patch can be applied.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Graph produced by this patch.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact semantic and physical no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the exact inverse patch without reinterpreting XML.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Replay the patch onto `package`, which must hold the captured source
    /// graph, and publish the result as a new package; `package` itself is
    /// unchanged.
    ///
    /// An exact no-op publishes `package` itself, so a signed package keeps
    /// its signature; a changed patch is refused for a signed package (ADR
    /// 0006): remove the signature explicitly first, with
    /// [`OpcPackage::unsign`] or the format facade's `unsign`.
    /// # Errors
    ///
    /// Returns [`crate::Error::Signed`] for a change to a signed package,
    /// [`crate::Error::PatchConflict`] when the recorded base differs, and
    /// an error when input
    /// violates OOXML constraints, exceeds a configured bound, or an
    /// underlying XML or package operation fails.
    pub fn apply(&self, package: &OpcPackage) -> Result<Commit> {
        let current = Snapshot::load_scoped(package, &self.before.scope_names())?;
        if !self.before.same_source(&current) {
            return Err(super::snapshot::source_mismatch());
        }
        if self.is_noop() {
            return Ok(Commit::new(
                package.clone(),
                self.clone(),
                Diagnostics::UNCHANGED,
            ));
        }
        refuse_signed_change(package)?;

        let mut candidate = package.clone();
        package::apply_items(&mut candidate, &self.before, self.after.items())?;
        let resulting = Snapshot::load_scoped(&candidate, &self.after.scope_names())?;
        if !self.after.same_source(&resulting) {
            return Err(crate::Error::Internal(
                "custom XML candidate does not match the prepared target",
            ));
        }
        Ok(Commit::new(candidate, self.clone(), Diagnostics::CHANGED))
    }
}

/// A published Custom XML change: the new package, its patch and
/// diagnostics. Read the resulting graph with
/// [`Snapshot::load`]`(commit.package())`.
pub type Commit = litchi_core::edit::Commit<OpcPackage, Patch>;

/// What a Custom XML commit changed.
pub use litchi_core::edit::Diagnostics;
