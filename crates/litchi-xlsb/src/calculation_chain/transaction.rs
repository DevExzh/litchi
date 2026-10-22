//! Detached source-bound Calculation Chain edits.

use super::patch::{Commit, Patch};
use super::snapshot::Snapshot;
use crate::package::error::{Error, Result};
use litchi_core::edit::Diagnostics;

crate::part_edit::limits_options!(super::SemanticLimits, "Calculation Chain");

/// A staged Calculation Chain edit of one package. Dropping it discards it.
#[derive(Clone, Debug)]
pub struct Edit {
    source: crate::Package,
    before: Snapshot,
    remove: bool,
}

impl Edit {
    pub(crate) fn new(source: crate::Package, before: Snapshot) -> Self {
        Self {
            source,
            remove: !before.is_present(),
            before,
        }
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.remove = !self.before.is_present();
    }

    /// Immutable source snapshot used by stale-source checks.
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the currently staged opaque owner, or `None` when removal is
    /// staged. The bytes remain source-backed and are never decoded.
    pub fn part(&self) -> Option<&super::ChainPart> {
        if self.remove {
            None
        } else {
            self.before.part()
        }
    }

    /// Stage removal of the optional Calculation Chain owner.
    pub fn remove(&mut self) -> Result<bool> {
        if !self.before.is_present() || self.remove {
            return Ok(false);
        }
        self.remove = true;
        Ok(true)
    }

    /// Validate the staged removal and publish it to a new package. An
    /// unchanged edit returns the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed edit of a signed workbook, the
    /// limit errors and the workbook validation errors of the result.
    pub fn commit(self) -> Result<Commit> {
        if !self.remove || !self.before.is_present() {
            let before = self.before;
            return Ok(Commit::new(
                self.source,
                Patch::new(before.clone(), before),
                Diagnostics::UNCHANGED,
            ));
        }
        if self.before.package().is_signed()
            || self.before.package().requires_signature_edit_policy()
        {
            return Err(Error::Signed);
        }

        let mut candidate = self.before.package().as_ref().clone();
        super::patch::remove_owner(&mut candidate, self.before.source(), self.before.limits())?;
        let after = Snapshot::read_with_limits(&candidate, self.before.limits())?;
        if after.is_present() {
            return Err(Error::InvalidFormat(
                "Calculation Chain removal readback still has an owner",
            ));
        }
        self.source
            .apply_calculation_chain(&Patch::new(self.before, after))
    }
}
