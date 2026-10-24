//! Source-bound package ownership for rich-value refresh interval metadata.

use std::sync::Arc;

use litchi_opc::{OpcPackage, OwnedRelationships, PackURI, Part as OpcPart, SourceLimits};

use crate::error::{Result, invalid};

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
    pub fn load(package: &OpcPackage, limits: &SourceLimits) -> Result<Self> {
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
        let inspection = codec::inspect(owner.blob(), limits)?;
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
        let relationships = package.source_relationships(&part_name, package.read_limits())?;
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
        return Err(super::limit(
            LimitKind::Relationships,
            values.len().checked_add(1).unwrap_or(usize::MAX),
            MAX_INCOMING_RELATIONSHIPS,
        ));
    }
    values.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "rich-value incoming relationship closure",
        source,
    })?;
    values.push(value);
    Ok(())
}

crate::record::record_family!("rich-value refresh-interval");

crate::record::limit_kinds!("rich-value refresh-interval", RichValueRefresh {
    XmlBytes => "rich-value refresh XML bytes",
    OutputBytes => "rich-value refresh output bytes",
    StringBytes => "rich-value refresh string bytes",
    Types => "rich-value types",
    Intervals => "rich-value refresh intervals",
    ExtensionLocations => "rich-value refresh extension locations",
    Relationships => "rich-value refresh incoming relationships",
    Depth => "rich-value XML depth",
});

/// A staged, failure-atomic edit of the rich-value types owner's refresh
/// intervals of one [`crate::Package`]. Dropping it discards it;
/// [`Self::commit`] publishes a new package.
#[derive(Debug, Clone)]
pub struct Edit {
    target: crate::record::Source,
    before: Option<Snapshot>,
    draft: Option<Vec<TypeRefreshIntervals>>,
}

impl Edit {
    /// Start a source-bound edit.
    fn start(target: crate::record::Source) -> Result<Self> {
        Ok(Self {
            target,
            before: None,
            draft: None,
        })
    }

    /// Read and clone the owner only when the edit first needs its staged
    /// semantic values. A failed read leaves the edit unprepared and can be
    /// retried without retaining a partial draft.
    fn ensure_prepared(&mut self) -> Result<()> {
        if self.before.is_some() {
            return Ok(());
        }
        let limits = self.target.read_limits();
        let before = Snapshot::load(&self.target, &limits)?;
        let draft = before.types.clone();
        self.draft = Some(draft);
        self.before = Some(before);
        Ok(())
    }

    fn snapshot(&self) -> Result<&Snapshot> {
        self.before
            .as_ref()
            .ok_or_else(|| invalid("rich-value refresh edit is not prepared"))
    }

    fn staged(&self) -> Result<&[TypeRefreshIntervals]> {
        self.draft
            .as_deref()
            .ok_or_else(|| invalid("rich-value refresh edit is not prepared"))
    }

    /// Exact immutable source captured when the edit started.
    #[must_use = "use the source snapshot or handle the error"]
    pub fn before(&mut self) -> Result<&Snapshot> {
        self.ensure_prepared()?;
        self.snapshot()
    }

    /// Currently staged type metadata.
    #[must_use = "use the staged type metadata or handle the error"]
    pub fn types(&mut self) -> Result<&[TypeRefreshIntervals]> {
        self.ensure_prepared()?;
        self.staged()
    }

    /// Replace one type's optional refresh collection.
    pub fn set(&mut self, type_index: usize, intervals: Option<RefreshIntervals>) -> Result<bool> {
        self.ensure_prepared()?;
        let value = self
            .draft
            .as_mut()
            .ok_or_else(|| invalid("rich-value refresh edit is not prepared"))?
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
        self.ensure_prepared()?;
        let current = self
            .draft
            .as_ref()
            .ok_or_else(|| invalid("rich-value refresh edit is not prepared"))?
            .get(type_index)
            .ok_or_else(|| invalid("rich-value type index is out of range"))?
            .intervals
            .clone();
        let mut values = current
            .map(|intervals| intervals.intervals().to_vec())
            .unwrap_or_default();
        if values.len() >= MAX_INTERVALS {
            return Err(super::limit(
                LimitKind::Intervals,
                values.len().checked_add(1).unwrap_or(usize::MAX),
                MAX_INTERVALS,
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
        self.ensure_prepared()?;
        let current = self
            .draft
            .as_ref()
            .ok_or_else(|| invalid("rich-value refresh edit is not prepared"))?
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

    /// Discard every staged change.
    pub fn reset(&mut self) -> Result<()> {
        self.ensure_prepared()?;
        let draft = self.snapshot()?.types.clone();
        self.draft = Some(draft);
        Ok(())
    }

    /// Validate and publish the staged owner edit as a new package. An
    /// unchanged edit returns the source package.
    ///
    /// Adding a payload requires an existing source `x:ext` owner.  The
    /// owner URI is opaque producer metadata, so a missing owner cannot be
    /// safely synthesized; that form returns an error and publishes nothing.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`], [`Error::PatchConflict`] and the
    /// publication's validation errors.
    pub fn commit(self) -> Result<Commit> {
        self.publish().map(|(commit, _)| commit)
    }

    fn publish(mut self) -> Result<(Commit, Snapshot)> {
        self.ensure_prepared()?;
        let before = self
            .before
            .take()
            .ok_or_else(|| invalid("rich-value refresh edit is not prepared"))?;
        let draft = self
            .draft
            .take()
            .ok_or_else(|| invalid("rich-value refresh edit is not prepared"))?;
        if before.types == draft {
            let patch = Patch::new(before.clone(), before.clone());
            return Ok((
                Commit::new(self.target.into_package(), patch, Diagnostics::new(false)),
                before,
            ));
        }
        if self.target.is_signed() {
            return Err(Error::Signed);
        }
        let limits = self.target.read_limits();
        let current = Snapshot::load(&self.target, &limits)?;
        if !current.same_source(&before) {
            return Err(Error::PatchConflict {
                part: "rich-value refreshIntervals source closure".to_owned(),
            });
        }
        let source = before
            .source
            .part
            .as_ref()
            .ok_or_else(|| invalid("rich-value types part is required for a changed edit"))?;
        let inspection = codec::inspect(source.bytes.as_slice(), &limits)?;
        let output = codec::rewrite(source.bytes.as_slice(), &inspection, &draft, &limits)?;
        let mut candidate = OpcPackage::clone(&self.target);
        candidate
            .get_part_mut(&source.name)?
            .set_blob_shared(Arc::new(output));
        let snapshot = Snapshot::load(&candidate, &limits)?;
        if snapshot.types != draft || !snapshot.same_closure(&before) {
            return Err(invalid(
                "rich-value refresh publication changed staged metadata or its relationship closure",
            ));
        }
        let patch = Patch::new(before, snapshot.clone());
        let package = self.target.publish(candidate)?;
        Ok((
            Commit::new(package, patch, Diagnostics::new(true)),
            snapshot,
        ))
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
    pub fn is_noop(&self) -> bool {
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
    pub(crate) fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let limits = target.read_limits();
        let current = Snapshot::load(target, &limits)?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: "rich-value refreshIntervals source closure".to_owned(),
            });
        }
        if self.is_noop() {
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
        let resulting = Snapshot::load(&candidate, &limits)?;
        if !resulting.same_source(&self.after) {
            return Err(Error::PatchConflict {
                part: "rich-value refreshIntervals patch result".to_owned(),
            });
        }
        *target = candidate;
        Ok(())
    }
}

impl crate::record::Record for Patch {
    fn is_noop(&self) -> bool {
        Self::is_noop(self)
    }

    fn apply_to(&self, package: &mut OpcPackage) -> Result<()> {
        self.apply(package)
    }
}

impl crate::Package {
    /// Start an edit of the rich-value types' refresh intervals.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused, and the validation errors of the rich-value types owner.
    pub fn edit_rich_value_refresh(&self) -> Result<Edit> {
        Edit::start(crate::record::Source::new(self, "edit_rich_value_refresh")?)
    }

    /// Apply an exact rich-value refresh-interval patch to a new package;
    /// this one is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_rich_value_refresh(&self, patch: &Patch) -> Result<Commit> {
        crate::record::replay(self, patch, "apply_rich_value_refresh")
    }
}

/// Load rich-value refresh metadata from an OPC package.
pub fn load(package: &OpcPackage, limits: &SourceLimits) -> Result<Snapshot> {
    Snapshot::load(package, limits)
}

#[cfg(test)]
mod lazy_tests {
    use super::*;
    use litchi_opc::BlobPart;

    #[test]
    fn edit_entry_defers_rich_value_type_decode() {
        let package = crate::Package::new().expect("minimal XLSX");
        let mut edit = package.edit_rich_value_refresh().expect("rich-value edit");
        assert!(edit.before.is_none());
        assert!(edit.draft.is_none());
        assert!(edit.types().expect("first type read").is_empty());
        assert!(edit.before.is_some());
    }

    #[test]
    fn failed_rich_value_decode_is_retryable_without_partial_state() {
        let source = crate::Package::new().expect("minimal XLSX");
        let mut package = source.opc().clone();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/richValueTypes.xml").expect("types URI"),
            TYPES_CONTENT_TYPE.into(),
            b"<rv:types".to_vec(),
        )));
        let package = crate::Package::from_opc(package).expect("defer rich-value XML");
        let mut edit = package.edit_rich_value_refresh().expect("rich-value edit");
        assert!(edit.types().is_err());
        assert!(edit.before.is_none());
        assert!(edit.types().is_err());
        assert!(edit.before.is_none());
    }
}
