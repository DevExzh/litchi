//! Reversible source-checked patches for the workbook external-link owner.

use litchi_opc::OpcPackage;

use super::package::apply_entries;
use super::snapshot::Snapshot;
use crate::error::Result;

/// A reversible replacement of the workbook's external-link graph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source context required before applying this patch.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Source context produced by this patch.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact source no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply atomically after checking the workbook owner and every owned
    /// external-link part.
    pub(crate) fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load(target)?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: self.before.workbook_part_name().to_owned(),
            });
        }
        if self.is_noop() {
            return Ok(());
        }

        let mut candidate = target.clone();
        apply_entries(
            &mut candidate,
            self.before.entries(),
            self.after.entries(),
            self.after.conformance(),
        )?;
        let resulting = Snapshot::load(&candidate)?;
        if !resulting.same_source(&self.after) || resulting.entries() != self.after.entries() {
            return Err(Error::PatchConflict {
                part: self.after.workbook_part_name().to_owned(),
            });
        }
        *target = candidate;
        Ok(())
    }
}

crate::record::record_family!("external-link");

crate::record::no_limits!("external-link");

impl crate::record::Record for Patch {
    fn is_noop(&self) -> bool {
        Self::is_noop(self)
    }

    fn apply_to(&self, package: &mut OpcPackage) -> Result<()> {
        self.apply(package)
    }
}
