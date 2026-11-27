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
    limits: super::SemanticLimits,
    before: Option<Snapshot>,
    remove: bool,
}

impl Edit {
    pub(crate) fn new(source: crate::Package, limits: super::SemanticLimits) -> Self {
        Self {
            source,
            limits,
            before: None,
            remove: false,
        }
    }

    fn ensure_prepared(&mut self) -> Result<&Snapshot> {
        if self.before.is_none() {
            self.before = Some(Snapshot::read_with_limits(
                self.source.opc_package(),
                self.limits,
            )?);
        }
        self.before.as_ref().ok_or(Error::InvalidFormat(
            "Calculation Chain edit is not prepared",
        ))
    }

    /// Discard every staged change.
    pub fn reset(&mut self) -> Result<()> {
        self.remove = !self.ensure_prepared()?.is_present();
        Ok(())
    }

    /// Immutable source snapshot used by stale-source checks.
    pub fn before(&mut self) -> Result<&Snapshot> {
        self.ensure_prepared()
    }

    /// Borrow the currently staged opaque owner, or `None` when removal is
    /// staged. The bytes remain source-backed and are never decoded.
    pub fn part(&mut self) -> Result<Option<&super::ChainPart>> {
        let remove = self.remove;
        let before = self.ensure_prepared()?;
        if remove { Ok(None) } else { Ok(before.part()) }
    }

    /// Stage removal of the optional Calculation Chain owner.
    pub fn remove(&mut self) -> Result<bool> {
        if !self.ensure_prepared()?.is_present() || self.remove {
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
    pub fn commit(mut self) -> Result<Commit> {
        self.ensure_prepared()?;
        let before = self.before.take().ok_or(Error::InvalidFormat(
            "Calculation Chain edit is not prepared",
        ))?;
        if !self.remove || !before.is_present() {
            return Ok(Commit::new(
                self.source,
                Patch::new(before.clone(), before),
                Diagnostics::UNCHANGED,
            ));
        }
        if before.package().is_signed() || before.package().requires_signature_edit_policy() {
            return Err(Error::Signed);
        }

        let mut candidate = before.package().as_ref().clone();
        super::patch::remove_owner(&mut candidate, before.source(), before.limits())?;
        let after = Snapshot::read_with_limits(&candidate, before.limits())?;
        if after.is_present() {
            return Err(Error::InvalidFormat(
                "Calculation Chain removal readback still has an owner",
            ));
        }
        self.source
            .apply_calculation_chain(&Patch::new(before, after))
    }
}
