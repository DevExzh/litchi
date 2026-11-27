//! Exact relationship-entry transport under source and namespace guards.
use super::super::metadata_fragments::{Fragment, FragmentEdit, FragmentPlan, Fragments, Kind};
use super::OwnedRelationships;
use crate::{PackURI, Result, SourceLimits};
use std::sync::Arc;

/// Bounded lexical entries decoded from one relationship member.
///
/// Decoding grants no package replacement authority. Fragment edits must still
/// match every touched destination entry and its inherited XML context.
#[derive(Debug)]
pub struct RelationshipFragments {
    owner: PackURI,
    inner: Fragments,
}

/// Borrowed exact entry, including its following-entry and trailing-byte anchor.
#[derive(Clone, Copy, Debug)]
pub struct RelationshipFragment<'a> {
    inner: Fragment<'a>,
    owner: &'a PackURI,
}
impl<'a> RelationshipFragment<'a> {
    /// Exact relationship ID within the owner.
    pub fn id(&self) -> &'a str {
        self.inner.id()
    }
    /// Exact element range in the catalog's original XML, including its BOM offset.
    pub fn byte_range(&self) -> std::ops::Range<usize> {
        self.inner.byte_range()
    }
    /// Exact element bytes, including attributes, paired tags and child markup.
    pub fn bytes(&self) -> &'a [u8] {
        self.inner.bytes()
    }
    /// Source relationship owner, or `/` for package relationships.
    pub fn owner(&self) -> &'a PackURI {
        self.owner
    }
    /// Next relationship in source order, if present.
    pub fn following_id(&self) -> Option<&'a str> {
        self.inner.following_id()
    }
    /// Exact bytes between this element and its following entry or root close.
    pub fn trailing_bytes(&self) -> &'a [u8] {
        self.inner.trailing_bytes()
    }
}
impl RelationshipFragments {
    /// Decode untrusted XML under byte, event, graph and structural bounds.
    /// Duplicate IDs and repaired/nonstandard lexical roots are not editable.
    pub fn from_bytes(owner: PackURI, bytes: Arc<Vec<u8>>, limits: SourceLimits) -> Result<Self> {
        let inner = Fragments::from_bytes(Kind::Relationships(owner.clone()), bytes, limits)?;
        Ok(Self { owner, inner })
    }
    /// Exact source-order entry views, without copying their XML.
    pub fn iter(&self) -> impl ExactSizeIterator<Item = RelationshipFragment<'_>> {
        self.inner.iter().map(|inner| RelationshipFragment {
            inner,
            owner: &self.owner,
        })
    }
    /// Resolve one exact relationship ID.
    pub fn get(&self, id: &str) -> Option<RelationshipFragment<'_>> {
        self.inner.get(id).map(|inner| RelationshipFragment {
            inner,
            owner: &self.owner,
        })
    }
}
/// One guarded entry delta. Replacement IDs and owners must agree.
#[derive(Clone, Copy, Debug)]
pub struct RelationshipFragmentEdit<'a> {
    /// Exact expected destination element, or absence for insertion.
    pub before: Option<RelationshipFragment<'a>>,
    /// Exact replacement element and anchors, or absence for removal.
    pub after: Option<RelationshipFragment<'a>>,
}
/// Source-bound batched lexical edit with exact byte/count admission.
///
/// The event bound is conservative: removing entries may merge adjacent text
/// events. No replacement XML buffer is allocated until materialization.
#[derive(Debug)]
pub struct RelationshipFragmentPlan<'a> {
    source: &'a OwnedRelationships,
    inner: FragmentPlan<'a>,
}
impl OwnedRelationships {
    /// Capture bounded exact entries for a durable metadata delta.
    pub fn fragments(&self, limits: SourceLimits) -> Result<RelationshipFragments> {
        let limits = self.xml.limits().min(limits);
        RelationshipFragments::from_bytes(self.owner.clone(), self.bytes_arc(), limits)
    }
    /// Plan guarded entry transport while retaining unrelated destination XML.
    ///
    /// Touched entries and inherited XML context must match. Insertions prove
    /// neighboring IDs and exact gaps; replacements retain destination slots.
    /// Missing, stale, cyclic or overlapping anchors fail before output allocation.
    /// Package-wide aggregate reservations and dependency closure are caller-owned.
    pub fn plan_fragment_edits<'a>(
        &'a self,
        edits: &[RelationshipFragmentEdit<'a>],
        limits: SourceLimits,
    ) -> Result<RelationshipFragmentPlan<'a>> {
        let limits = self.xml.limits().min(limits);
        Kind::Relationships(self.owner.clone()).admit_edit_count(edits.len(), limits)?;
        let mut inner_edits = Vec::new();
        inner_edits
            .try_reserve_exact(edits.len())
            .map_err(|source| crate::OpcError::Allocation {
                resource: "OPC relationship fragment edits",
                source,
            })?;
        inner_edits.extend(edits.iter().map(|edit| FragmentEdit {
            before: edit.before.map(|value| value.inner),
            after: edit.after.map(|value| value.inner),
        }));
        let inner = FragmentPlan::new(
            &self.xml,
            Kind::Relationships(self.owner.clone()),
            &inner_edits,
            limits,
        )?;
        Ok(RelationshipFragmentPlan {
            source: self,
            inner,
        })
    }
}
impl RelationshipFragmentPlan<'_> {
    /// Exact final relationship-member byte length.
    pub const fn final_len(&self) -> usize {
        self.inner.final_len()
    }
    /// Exact final relationship count.
    pub const fn relationship_count(&self) -> usize {
        self.inner.entry_count()
    }
    /// Conservative XML event bound, including EOF.
    pub const fn event_upper_bound(&self) -> usize {
        self.inner.event_upper_bound()
    }
    /// Whether materialization returns the exact source image.
    pub const fn is_noop(&self) -> bool {
        self.inner.is_noop()
    }
    /// Whether the resulting relationship member must be published.
    pub fn member_present(&self) -> bool {
        self.source.member_present || self.relationship_count() != 0
    }
    /// Materialize and validate XML under the supplied limits, even for no-ops.
    pub fn materialize(self, limits: SourceLimits) -> Result<OwnedRelationships> {
        let member_present = self.member_present();
        let (xml, _) = self.inner.materialize(limits)?;
        Ok(OwnedRelationships {
            owner: self.source.owner.clone(),
            xml,
            member_present,
        })
    }
}
