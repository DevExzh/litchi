//! Source-bound package ownership for rich-value refresh interval metadata.

use std::sync::Arc;

use litchi_opc::{OpcPackage, OwnedRelationships, PackURI, Part as OpcPart};

use crate::error::{Error, Result, invalid};

use super::MAX_INTERVALS;
use super::codec;
use super::model::{RefreshInterval, RefreshIntervals, TypeRefreshIntervals, validate_intervals};

const TYPES_CONTENT_TYPE: &str = "application/vnd.ms-excel.rdRichValuetypes+xml";
const MAX_INCOMING_RELATIONSHIPS: usize = super::super::MAX_RELATIONSHIPS;

/// Immutable source-bound rich-value refresh metadata.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    types: Vec<TypeRefreshIntervals>,
    source: SourceState,
}

impl Snapshot {
    /// Load the unique `rdRichValuetypes` owner, if present.
    ///
    /// The loader captures only the owner part and its incoming/outgoing OPC
    /// relationship closure.  It never opens a target or executes a refresh.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        let mut owner = None;
        for part in package.iter_parts() {
            if part.content_type() == TYPES_CONTENT_TYPE {
                if owner.is_some() {
                    return Err(invalid("package contains multiple rich-value types parts"));
                }
                owner = Some(part);
            }
        }
        let Some(owner) = owner else {
            return Ok(Self {
                types: Vec::new(),
                source: SourceState::absent(),
            });
        };
        // Only the owner's payload is read, so only it is decoded (ADR 0030).
        let owner = package.get_part(owner.partname())?;
        let inspection = codec::inspect(owner.blob())?;
        let source = SourceState::capture(package, owner)?;
        Ok(Self {
            types: inspection
                .types
                .into_iter()
                .map(|entry| entry.model)
                .collect(),
            source,
        })
    }

    /// Alias emphasizing the source-bound read operation.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        Self::load(package)
    }

    /// Every `CT_RichValueType` in source order, including types with no
    /// refresh payload.  The slice index is the transaction selector.
    #[must_use]
    pub fn types(&self) -> &[TypeRefreshIntervals] {
        &self.types
    }

    /// The exact rich-value types part URI, if present.
    #[must_use]
    pub fn part_name(&self) -> Option<&PackURI> {
        self.source.part.as_ref().map(|part| &part.name)
    }

    /// Exact source bytes of the rich-value types part, if present.
    #[must_use]
    pub fn source_xml(&self) -> Option<&[u8]> {
        self.source.part.as_ref().map(|part| part.bytes.as_slice())
    }

    /// Shared ownership of the exact source bytes, useful for proving an
    /// exact semantic no-op did not allocate a replacement owner.
    #[must_use]
    pub fn source_arc(&self) -> Option<Arc<Vec<u8>>> {
        self.source
            .part
            .as_ref()
            .map(|part| Arc::clone(&part.bytes))
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.source == other.source
    }

    fn same_closure(&self, other: &Self) -> bool {
        self.source.same_closure(&other.source)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SourceState {
    part: Option<SourcePart>,
    incoming: Vec<RelationshipState>,
}

impl SourceState {
    fn absent() -> Self {
        Self {
            part: None,
            incoming: Vec::new(),
        }
    }

    fn capture(package: &OpcPackage, owner: &dyn OpcPart) -> Result<Self> {
        let part_name = owner.partname().clone();
        let relationships = package.source_relationships(&part_name)?;
        Ok(Self {
            part: Some(SourcePart {
                name: part_name.clone(),
                content_type: owner.content_type().to_owned(),
                bytes: owner.blob_arc(),
                relationships,
            }),
            incoming: capture_incoming(package, &part_name)?,
        })
    }

    fn same_closure(&self, other: &Self) -> bool {
        match (&self.part, &other.part) {
            (None, None) => self.incoming == other.incoming,
            (Some(left), Some(right)) => {
                left.name == right.name
                    && left.content_type == right.content_type
                    && left.relationships == right.relationships
                    && self.incoming == other.incoming
            },
            _ => false,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SourcePart {
    name: PackURI,
    content_type: String,
    bytes: Arc<Vec<u8>>,
    relationships: OwnedRelationships,
}

#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
struct RelationshipState {
    source: String,
    id: String,
    reltype: String,
    target: String,
    external: bool,
}

fn capture_incoming(package: &OpcPackage, target: &PackURI) -> Result<Vec<RelationshipState>> {
    let mut incoming = Vec::new();
    for relationship in package.rels().iter() {
        if !relationship.is_external() && relationship.target_partname()?.is_equivalent_to(target) {
            push_incoming(
                &mut incoming,
                RelationshipState {
                    source: "/".to_owned(),
                    id: relationship.r_id().to_owned(),
                    reltype: relationship.reltype().to_owned(),
                    target: relationship.target_ref().to_owned(),
                    external: relationship.is_external(),
                },
            )?;
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if !relationship.is_external()
                && relationship.target_partname()?.is_equivalent_to(target)
            {
                push_incoming(
                    &mut incoming,
                    RelationshipState {
                        source: part.partname().as_str().to_owned(),
                        id: relationship.r_id().to_owned(),
                        reltype: relationship.reltype().to_owned(),
                        target: relationship.target_ref().to_owned(),
                        external: relationship.is_external(),
                    },
                )?;
            }
        }
    }
    incoming.sort_unstable();
    Ok(incoming)
}

fn push_incoming(values: &mut Vec<RelationshipState>, value: RelationshipState) -> Result<()> {
    if values.len() >= MAX_INCOMING_RELATIONSHIPS {
        return Err(invalid(
            "rich-value incoming relationship count exceeds the limit",
        ));
    }
    values.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "rich-value incoming relationship closure",
        source,
    })?;
    values.push(value);
    Ok(())
}

/// Failure-atomic edits over one rich-value types owner.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    draft: Vec<TypeRefreshIntervals>,
}

impl<'a> Transaction<'a> {
    /// Start a source-bound transaction.
    pub fn new(target: &'a mut OpcPackage) -> Result<Self> {
        let before = Snapshot::load(target)?;
        Ok(Self {
            draft: before.types.clone(),
            target,
            before,
        })
    }

    /// Exact immutable source captured at transaction start.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Currently staged type metadata.
    #[must_use]
    pub fn types(&self) -> &[TypeRefreshIntervals] {
        &self.draft
    }

    /// Replace one type's optional refresh collection.
    pub fn set(&mut self, type_index: usize, intervals: Option<RefreshIntervals>) -> Result<bool> {
        let value = self
            .draft
            .get_mut(type_index)
            .ok_or_else(|| invalid("rich-value type index is out of range"))?;
        if value.intervals == intervals {
            return Ok(false);
        }
        if let Some(intervals) = intervals.as_ref() {
            validate_intervals(intervals.intervals())?;
        }
        value.intervals = intervals;
        Ok(true)
    }

    /// Append one interval, creating the optional collection when necessary.
    pub fn add(&mut self, type_index: usize, interval: RefreshInterval) -> Result<bool> {
        let current = self
            .draft
            .get(type_index)
            .ok_or_else(|| invalid("rich-value type index is out of range"))?
            .intervals
            .clone();
        let mut values = current
            .map(|intervals| intervals.intervals().to_vec())
            .unwrap_or_default();
        if values.len() >= MAX_INTERVALS {
            return Err(invalid(
                "rich-value refresh interval count exceeds the limit",
            ));
        }
        values.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "rich-value refresh interval edit",
            source,
        })?;
        values.push(interval);
        self.set(type_index, Some(RefreshIntervals::new(values)?))
    }

    /// Remove one interval by its stable source-order index.
    pub fn remove(
        &mut self,
        type_index: usize,
        interval_index: usize,
    ) -> Result<Option<RefreshInterval>> {
        let current = self
            .draft
            .get(type_index)
            .ok_or_else(|| invalid("rich-value type index is out of range"))?
            .intervals
            .clone();
        let Some(current) = current else {
            return Ok(None);
        };
        let mut values = current.into_intervals();
        if interval_index >= values.len() {
            return Ok(None);
        }
        let removed = values.remove(interval_index);
        let replacement = if values.is_empty() {
            None
        } else {
            Some(RefreshIntervals::new(values)?)
        };
        self.set(type_index, replacement)?;
        Ok(Some(removed))
    }

    /// Remove all refresh metadata from one type.
    pub fn clear(&mut self, type_index: usize) -> Result<bool> {
        self.set(type_index, None)
    }

    /// Whether staged typed metadata differs from the source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.types != self.draft
    }

    /// Validate and atomically publish the staged owner edit.
    ///
    /// Adding a payload requires an existing source `x:ext` owner.  The
    /// owner URI is opaque producer metadata, so a missing owner cannot be
    /// safely synthesized; that form returns an error without changing the
    /// package.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            return Ok(Commit::new(
                self.before.clone(),
                Patch::new(self.before.clone(), self.before.clone()),
                false,
            ));
        }
        if self.target.is_signed() {
            return Err(Error::Signed);
        }
        let current = Snapshot::load(self.target)?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: "rich-value refreshIntervals source closure".to_owned(),
            });
        }
        let source = self
            .before
            .source
            .part
            .as_ref()
            .ok_or_else(|| invalid("rich-value types part is required for a changed edit"))?;
        let inspection = codec::inspect(source.bytes.as_slice())?;
        let output = codec::rewrite(source.bytes.as_slice(), &inspection, &self.draft)?;
        let mut candidate = self.target.clone();
        candidate
            .get_part_mut(&source.name)?
            .set_blob_shared(Arc::new(output));
        let snapshot = Snapshot::load(&candidate)?;
        if snapshot.types != self.draft || !snapshot.same_closure(&self.before) {
            return Err(invalid(
                "rich-value refresh publication changed staged metadata or its relationship closure",
            ));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        *self.target = candidate;
        Ok(Commit::new(snapshot, patch, true))
    }
}

/// An exact source-checked reversible refresh metadata edit.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source state required before application.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Source state produced by application.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact source no-op.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the exact source-bound inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply atomically after checking the complete owner relationship closure.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load(target)?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: "rich-value refreshIntervals source closure".to_owned(),
            });
        }
        if self.is_empty() {
            return Ok(());
        }
        if target.is_signed() {
            return Err(Error::Signed);
        }
        let source = self
            .after
            .source
            .part
            .as_ref()
            .ok_or_else(|| invalid("changed refresh patch has no types part"))?;
        let mut candidate = target.clone();
        candidate
            .get_part_mut(&source.name)?
            .set_blob_shared(Arc::clone(&source.bytes));
        let resulting = Snapshot::load(&candidate)?;
        if !resulting.same_source(&self.after) {
            return Err(Error::PatchConflict {
                part: "rich-value refreshIntervals patch result".to_owned(),
            });
        }
        *target = candidate;
        Ok(())
    }
}

/// A successfully staged refresh metadata publication.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn new(snapshot: Snapshot, patch: Patch, changed: bool) -> Self {
        Self {
            snapshot,
            patch,
            changed,
        }
    }

    /// Whether the typed metadata changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.changed
    }

    /// Resulting source-bound snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Exact reversible patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }
}

/// Load rich-value refresh metadata from an OPC package.
pub fn load(package: &OpcPackage) -> Result<Snapshot> {
    Snapshot::load(package)
}

/// Start a source-bound rich-value refresh transaction.
pub fn edit(package: &mut OpcPackage) -> Result<Transaction<'_>> {
    Transaction::new(package)
}

/// Apply an exact source-bound rich-value refresh patch.
pub fn apply_patch(package: &mut OpcPackage, patch: &Patch) -> Result<()> {
    patch.apply(package)
}
