//! Reversible source-checked patches for worksheet OLE metadata.

use litchi_opc::OpcPackage;

use super::snapshot::Snapshot;
use super::validation;
use crate::error::Result;

/// A reversible replacement of the known OLE metadata spans in one worksheet.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
    /// The exact part-level change of a patch that added a collection with
    /// its relationships and parts.
    parts: Option<crate::part_edit::PartPatch>,
}

impl Patch {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self {
            before,
            after,
            parts: None,
        }
    }

    pub(crate) fn with_parts(
        before: Snapshot,
        after: Snapshot,
        parts: crate::part_edit::PartPatch,
    ) -> Self {
        Self {
            before,
            after,
            parts: Some(parts),
        }
    }

    /// Source context required before this patch can be applied.
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
            parts: self
                .parts
                .as_ref()
                .map(crate::part_edit::PartPatch::inverse),
        }
    }

    /// Apply atomically after checking the worksheet and opaque resources.
    pub(crate) fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load(target, self.before.worksheet())?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: self.before.worksheet().to_string(),
            });
        }
        if self.is_noop() {
            return Ok(());
        }
        if let Some(parts) = &self.parts {
            let candidate = parts.apply_graph(target)?;
            let resulting = Snapshot::load(&candidate, self.after.worksheet())?;
            if !resulting.same_source(&self.after) {
                return Err(Error::PatchConflict {
                    part: self.after.worksheet().to_string(),
                });
            }
            *target = candidate;
            return Ok(());
        }
        let mut candidate = target.clone();
        candidate
            .get_part_mut(self.after.worksheet())?
            .set_blob(self.after.source_xml().to_vec());
        validation::graph(&candidate, self.after.worksheet())?;
        let resulting = Snapshot::load(&candidate, self.after.worksheet())?;
        if !resulting.same_source(&self.after) {
            return Err(Error::PatchConflict {
                part: self.after.worksheet().to_string(),
            });
        }
        *target = candidate;
        Ok(())
    }
}

crate::record::record_family!("OLE-object");

crate::record::limit_kinds!("OLE-object", OleObjects {
    XmlBytes => "worksheet OLE XML bytes",
    Nodes => "worksheet OLE XML nodes",
    Depth => "worksheet OLE XML depth",
    Objects => "worksheet OLE objects",
    PayloadBytes => "worksheet OLE payload bytes",
    TotalPayloadBytes => "worksheet OLE total payload bytes",
    StringBytes => "worksheet OLE string bytes",
});

impl crate::record::Record for Patch {
    fn is_noop(&self) -> bool {
        Self::is_noop(self)
    }

    fn apply_to(&self, package: &mut OpcPackage) -> Result<()> {
        self.apply(package)
    }
}
