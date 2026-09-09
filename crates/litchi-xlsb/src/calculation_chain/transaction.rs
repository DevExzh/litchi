//! Detached source-bound Calculation Chain edits.

use super::patch::{Commit, Patch};
use super::snapshot::Snapshot;
use crate::package::error::{Error, Result};

/// A bounded draft detached from an immutable Calculation Chain snapshot.
#[derive(Clone, Debug)]
pub struct Transaction {
    before: Snapshot,
    remove: bool,
}

impl Transaction {
    pub(crate) fn new(before: Snapshot) -> Self {
        Self {
            remove: !before.is_present(),
            before,
        }
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

    /// Commit the detached removal into a reversible source-checked patch.
    pub fn commit(self) -> Result<Commit> {
        if !self.remove || !self.before.is_present() {
            let before = self.before;
            return Ok(Commit::new(Patch::new(before.clone(), before), false));
        }
        if self.before.package().is_signed()
            || self.before.package().requires_signature_edit_policy()
        {
            return Err(Error::Opc(
                litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
            ));
        }

        let mut candidate = self.before.package().as_ref().clone();
        super::patch::remove_owner(&mut candidate, self.before.source(), self.before.limits())?;
        let after = Snapshot::read_with_limits(&candidate, self.before.limits())?;
        if after.is_present() {
            return Err(Error::InvalidFormat(
                "Calculation Chain removal readback still has an owner".to_string(),
            ));
        }
        Ok(Commit::new(Patch::new(self.before, after), true))
    }
}
