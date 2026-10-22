//! Reversible, source-checked patches for workbook revision metadata.

use litchi_opc::OpcPackage;

use crate::error::Result;

use super::package::{replace_workbook_revisions, restore_snapshot_source};
use super::snapshot::Snapshot;

/// A reversible replacement of the complete workbook revision owner.
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

    /// Whether this patch preserves the exact source owner.
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

    /// Apply atomically after checking the complete revision source graph.
    ///
    /// An empty patch publishes nothing and keeps a signed target exactly as
    /// it was. Any other patch changes the target, so a signed target is
    /// refused with [`Error::Signed`]: remove its signature explicitly first
    /// ([`OpcPackage::unsign`]). The inverse never restores a signature.
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
        crate::signature_policy::ensure_unsigned(target)?;

        let mut candidate = target.clone();
        replace_workbook_revisions(
            &mut candidate,
            self.after.revisions(),
            self.after
                .conformance()
                .unwrap_or(super::model::RevisionConformance::Transitional),
        )?;
        restore_snapshot_source(&mut candidate, &self.after)?;
        let resulting = Snapshot::load(&candidate)?;
        if !resulting.same_source(&self.after) {
            return Err(Error::PatchConflict {
                part: self.after.workbook_part_name().to_owned(),
            });
        }
        *target = candidate;
        Ok(())
    }
}

crate::record::record_family!("revision");

crate::record::limit_kinds!("revision", Revisions {
    /// Encoded or expanded XML bytes in one revision part.
    PartBytes => "revision part bytes",
    /// Nested revision XML elements.
    XmlDepth => "revision XML depth",
    /// Revision XML elements in one part.
    XmlNodes => "revision XML nodes",
    /// UTF-8 text and attribute bytes.
    TextBytes => "revision text bytes",
    /// Users in the revision catalog.
    Users => "revision users",
    /// Headers in the revision catalog.
    Headers => "revision headers",
    /// Revision log parts.
    Logs => "revision logs",
    /// Records in one revision log.
    Records => "revision records per log",
    /// Records across all revision logs.
    TotalRecords => "total revision records",
    /// Sheet identifiers in revision headers.
    SheetIds => "revision sheet identifiers",
});

impl crate::record::Record for Patch {
    fn is_noop(&self) -> bool {
        Self::is_noop(self)
    }

    fn apply_to(&self, package: &mut OpcPackage) -> Result<()> {
        self.apply(package)
    }
}
