//! Reversible source-checked patches for the workbook external-link owner.

use litchi_opc::OpcPackage;

use super::snapshot::Snapshot;
use crate::error::{Error as XlsxError, Result};

/// A reversible replacement of the workbook's external-link graph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
    image: crate::part_edit::PartPatch,
}

impl Patch {
    pub(crate) fn new(
        before: Snapshot,
        after: Snapshot,
        image: crate::part_edit::PartPatch,
    ) -> Self {
        Self {
            before,
            after,
            image,
        }
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
            image: self.image.inverse(),
        }
    }

    /// Apply atomically after checking the workbook owner and every owned
    /// external-link part.
    pub(crate) fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let limits = target.read_limits();
        let current = Snapshot::load(target, &limits)?;
        if !current.same_source(&self.before) {
            return Err(XlsxError::PatchConflict {
                part: self.before.workbook_part_name().to_owned(),
            });
        }
        if self.is_noop() {
            return Ok(());
        }

        let candidate = self.image.apply_graph(target)?;
        let resulting = Snapshot::load(&candidate, &limits)?;
        if !resulting.same_source(&self.after) || resulting.entries() != self.after.entries() {
            return Err(XlsxError::PatchConflict {
                part: self.after.workbook_part_name().to_owned(),
            });
        }
        *target = candidate;
        Ok(())
    }
}

crate::record::record_family!("external-link");

crate::record::limit_kinds!("external-link", ExternalLinks {
    XmlBytes => "external-link XML bytes",
    Events => "external-link XML events",
    Depth => "external-link XML depth",
    TargetUri => "external-link target URI bytes",
    SheetNames => "external-link sheet names",
    DefinedNames => "external-link defined names",
    CachedSheets => "external-link cached sheets",
    CachedRows => "external-link cached rows",
    CachedCells => "external-link cached cells",
    Items => "external-link items",
    MatrixCells => "external-link matrix cells",
    CacheTextBytes => "external-link cache text bytes",
    RelationshipIdBytes => "external-link relationship ID bytes",
    AlternateUrlMetadataBytes => "external-link alternate URL metadata bytes",
    CellColumn => "external-link cell column",
    CellRow => "external-link cell row",
    ColumnLetters => "external-link column letters",
});

impl crate::record::Record for Patch {
    fn is_noop(&self) -> bool {
        Self::is_noop(self)
    }

    fn apply_to(&self, package: &mut OpcPackage) -> Result<()> {
        self.apply(package)
    }
}
