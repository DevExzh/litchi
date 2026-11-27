//! Exact reversible worksheet page-break patches.

use litchi_opc::OpcPackage;

use super::Snapshot;
use crate::error::{Error, Result, invalid};

/// Exact source-checked page-break worksheet replacement.
#[derive(Clone, Debug)]
pub(crate) struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(super) const fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Required source state.
    #[must_use]
    pub(crate) const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Exact state produced by application.
    #[must_use]
    pub(crate) const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch preserves its source byte-for-byte.
    #[must_use]
    pub(crate) fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the exact source-bound inverse.
    #[must_use]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply after checking the complete retained workbook/worksheet owner.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source bytes, signed packages, invalid OPC
    /// state, or a target that fails semantic readback.
    pub(crate) fn apply(&self, package: &mut OpcPackage) -> Result<()> {
        self.before.check_execution()?;
        self.after.check_execution()?;
        if !self.before.matches_current_source(package) {
            return Err(Error::PatchConflict {
                part: self.before.worksheet_part_name().to_string(),
            });
        }
        if self.is_empty() {
            self.before.check_execution()?;
            self.after.check_execution()?;
            return Ok(());
        }
        if package.is_signed() {
            return Err(Error::Signed);
        }
        let mut candidate = package.clone();
        candidate
            .get_part_mut(self.before.worksheet_part_name())?
            .set_blob_shared(self.after.source_arc()?);
        let resulting = Snapshot::load(&candidate, self.after.sheet_position())?;
        self.after.check_execution()?;
        if !resulting.same_source(&self.after)
            || resulting.page_breaks() != self.after.page_breaks()
        {
            return Err(invalid(
                "page-break patch readback did not match its target",
            ));
        }
        *package = candidate;
        Ok(())
    }
}
