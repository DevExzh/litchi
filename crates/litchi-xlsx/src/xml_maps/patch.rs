//! Reversible source-checked patches for workbook Custom XML Maps.

use crate::error::Result;
use litchi_opc::OpcPackage;

use super::invalid;
use super::snapshot::Snapshot;

/// A reversible replacement of the workbook's Custom XML Maps graph.
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

    /// Apply atomically after checking the workbook and owned-part source.
    ///
    /// An empty patch publishes nothing and keeps a signed target exactly as
    /// it was. Any other patch changes the target, so a signed target is
    /// refused with a boxed [`crate::Error::Signed`]: remove its signature
    /// explicitly first ([`OpcPackage::unsign`]). The inverse never restores
    /// a signature.
    pub(crate) fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load(target)?;
        if !current.same_source(&self.before) {
            return Err(invalid("custom XML maps patch source is stale"));
        }
        if self.is_noop() {
            return Ok(());
        }
        crate::signature_policy::ensure_unsigned(target)?;

        let mut candidate = target.clone();
        self.after.restore_into(&mut candidate)?;
        let resulting = Snapshot::load(&candidate)?;
        if !resulting.same_source(&self.after)
            || resulting.info() != self.after.info()
            || resulting.conformance() != self.after.conformance()
        {
            return Err(invalid(
                "custom XML maps patch publication changed its source",
            ));
        }
        *target = candidate;
        Ok(())
    }
}

crate::record::record_family!("Custom XML Maps");

crate::record::limit_kinds!("Custom XML Maps", XmlMaps {
    /// Bytes of the Custom XML Maps part.
    PartBytes => "bytes of the Custom XML Maps part",
    /// Schemas of the Custom XML Maps part.
    Schemas => "schemas of the Custom XML Maps part",
    /// Maps of the Custom XML Maps part.
    Maps => "maps of the Custom XML Maps part",
});

impl crate::record::Record for Patch {
    fn is_noop(&self) -> bool {
        Self::is_noop(self)
    }

    fn apply_to(&self, package: &mut OpcPackage) -> Result<()> {
        self.apply(package)
    }
}
