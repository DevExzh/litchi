//! Source-checked, reversible classic-comments package patches.

use litchi_opc::{OpcPackage, PackURI};

use super::package::{remove_from_worksheet, replace_on_worksheet, validate_graph};
use super::snapshot::Snapshot;
use crate::error::{Result, invalid};

/// A reversible source-checked replacement of one worksheet comments graph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
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

    /// Whether applying this patch is an exact source no-op.
    #[must_use]
    pub fn is_empty(&self) -> bool {
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

    /// Apply this patch atomically after checking the complete source owner.
    ///
    /// The package is cloned before any relationship or part operation. A
    /// source conflict, allocation failure, malformed graph, or identity
    /// mismatch therefore leaves the caller's package untouched. The captured
    /// comments bytes are restored after the existing package writer has
    /// planned the relationship operation, preserving the exact patch source.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load(target, self.before.worksheet())?;
        if !current.same_source(&self.before) {
            return Err(crate::error::Error::PatchConflict {
                part: self.before.worksheet().to_string(),
            });
        }
        if self.is_empty() {
            return Ok(());
        }

        let mut candidate = target.clone();
        if let Some(part) = self.after.part() {
            let stored =
                replace_on_worksheet(&mut candidate, self.after.worksheet(), &part.comments)?;
            if stored.relationship_id != part.relationship_id || stored.part_name != part.part_name
            {
                return Err(crate::error::Error::PatchConflict {
                    part: self.after.worksheet().to_string(),
                });
            }
            let relationship = candidate
                .get_part(self.after.worksheet())?
                .rels()
                .get(&stored.relationship_id)
                .ok_or_else(|| invalid("classic comments relationship was not published"))?;
            if Some(relationship.reltype()) != self.after.relationship_type() {
                return Err(crate::error::Error::PatchConflict {
                    part: self.after.worksheet().to_string(),
                });
            }
            let name = PackURI::new(&stored.part_name).map_err(invalid)?;
            let source = self
                .after
                .source_xml()
                .ok_or_else(|| invalid("classic comments patch is missing its after source"))?;
            let resource = candidate.get_part_mut(&name)?;
            resource.set_blob(source.to_vec());
        } else {
            remove_from_worksheet(&mut candidate, self.after.worksheet())?;
        }
        validate_graph(&candidate)?;
        let resulting = Snapshot::load(&candidate, self.after.worksheet())?;
        if !resulting.same_source(&self.after) {
            return Err(crate::error::Error::PatchConflict {
                part: self.after.worksheet().to_string(),
            });
        }
        *target = candidate;
        Ok(())
    }
}

/// Successful publication of one classic-comments transaction.
#[derive(Debug)]
pub struct Commit {
    pub(crate) snapshot: Snapshot,
    pub(crate) patch: Patch,
    pub(crate) changed: bool,
}

impl Commit {
    /// Whether the transaction changed the worksheet comments graph.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.changed
    }

    /// Resulting immutable source snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible source-checked package patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the publication into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}
