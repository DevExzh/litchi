#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
//! Source-checked transactions over a DOCX chart ownership graph.

use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI, Part, TargetMode};

use super::codec::{invalid, ownership, validate_graph_value};
use super::model::Graph;
use super::package;
use crate::{Error, Result};

/// Stable fingerprint of the captured document/chart source closure.
pub type Revision = u64;

/// An immutable source-bound snapshot of one DOCX chart graph.
///
/// The semantic [`Graph`] contains the exact chart, chart-style, color-style,
/// and embedded-workbook payloads. The private source image additionally
/// retains the document owner and every owned part's complete OPC relationship
/// set. This lets package edits reject stale owners without copying unrelated
/// package parts.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    document_name: PackURI,
    graph: Arc<Graph>,
    source: Source,
}

impl Snapshot {
    /// Load and validate the complete chart graph owned by `document_name`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn load(package: &OpcPackage, document_name: &PackURI) -> Result<Self> {
        let graph = package::load(package, document_name)?;
        let document = package.get_part(document_name)?;
        let document_name = document.partname().clone();
        let source = Source::capture(package, &document_name, &graph)?;
        Ok(Self {
            document_name,
            graph: Arc::new(graph),
            source,
        })
    }

    /// Alias emphasizing that the snapshot is bound to package source bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn read(package: &OpcPackage, document_name: &PackURI) -> Result<Self> {
        Self::load(package, document_name)
    }

    /// Borrow the validated semantic chart graph.
    #[must_use]
    pub fn graph(&self) -> &Graph {
        self.graph.as_ref()
    }

    /// Borrow charts in their document anchor order.
    #[must_use]
    pub fn charts(&self) -> &[super::model::Resource] {
        &self.graph.charts
    }

    /// Return the physical DOCX main-document part owned by this snapshot.
    #[must_use]
    pub const fn document_name(&self) -> &PackURI {
        &self.document_name
    }

    /// Alias for callers that name the owner as a part.
    #[must_use]
    pub const fn document_part_name(&self) -> &PackURI {
        self.document_name()
    }

    /// Borrow the exact source bytes of the document part.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source.document.data.as_slice()
    }

    /// Return the source fingerprint used by optimistic stale checks.
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.source.revision
    }

    /// Chart snapshots are always tied to an OPC source graph.
    #[must_use]
    pub const fn is_source_bound(&self) -> bool {
        true
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.document_name == other.document_name
            && self.graph == other.graph
            && self.source == other.source
    }
}

/// A failure-atomic edit over one existing DOCX chart ownership graph.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    draft: Graph,
}

impl<'a> Transaction<'a> {
    /// Capture the current chart graph and begin a package-bound transaction.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn new(target: &'a mut OpcPackage, document_name: &PackURI) -> Result<Self> {
        let before = Snapshot::load(target, document_name)?;
        Ok(Self {
            draft: before.graph().clone(),
            target,
            before,
        })
    }

    /// Alias for [`Self::new`] at call sites that prefer a verb.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn begin(target: &'a mut OpcPackage, document_name: &PackURI) -> Result<Self> {
        Self::new(target, document_name)
    }

    /// Borrow the immutable source snapshot used for conflict checks.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Alias for [`Self::before`].
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        self.before()
    }

    /// Borrow the currently staged semantic chart graph.
    #[must_use]
    pub const fn graph(&self) -> &Graph {
        &self.draft
    }

    /// Borrow the currently staged charts in document anchor order.
    #[must_use]
    pub fn charts(&self) -> &[super::model::Resource] {
        &self.draft.charts
    }

    /// Whether staged graph content differs from the captured source graph.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.draft != *self.before.graph
    }

    /// Replace the complete semantic graph after bounded graph validation.
    ///
    /// The package service still enforces stable ownership at publication;
    /// changing chart part identities or relationship ownership is rejected
    /// before the target package is touched.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn replace_graph(&mut self, graph: Graph) -> Result<bool> {
        validate_graph_value(&graph)?;
        if self.draft == graph {
            return Ok(false);
        }
        self.draft = graph;
        Ok(true)
    }

    /// Alias for [`Self::replace_graph`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn replace(&mut self, graph: Graph) -> Result<bool> {
        self.replace_graph(graph)
    }

    /// Apply a checked, failure-atomic mutation to the staged graph.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn edit_graph(&mut self, edit: impl FnOnce(&mut Graph) -> Result<()>) -> Result<&mut Self> {
        let mut candidate = self.draft.clone();
        edit(&mut candidate)?;
        validate_graph_value(&candidate)?;
        self.draft = candidate;
        Ok(self)
    }

    /// Alias for [`Self::edit_graph`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn edit(&mut self, edit: impl FnOnce(&mut Graph) -> Result<()>) -> Result<&mut Self> {
        self.edit_graph(edit)
    }

    /// Validate and publish this graph as a reversible source-checked patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn commit(self) -> Result<Commit> {
        let current = Snapshot::load(self.target, self.before.document_name())?;
        if !current.same_source(&self.before) {
            return Err(invalid("chart transaction source is stale"));
        }
        if self.draft == *self.before.graph {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }

        validate_graph_value(&self.draft)?;
        let mut candidate = self.target.clone();
        package::store(&mut candidate, self.before.document_name(), &self.draft)?;
        let snapshot = Snapshot::load(&candidate, self.before.document_name())?;
        if snapshot.graph() != &self.draft {
            return Err(invalid("chart publication changed the staged graph"));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        *self.target = candidate;
        Ok(Commit::new(snapshot, patch, true))
    }

    /// Discard staged changes and recover the source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.before
    }
}

/// A successful chart publication and its reversible source patch.
#[derive(Clone, Debug)]
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

    /// Whether the transaction changed any chart-owned source bytes.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Alias for [`Self::changed`].
    #[must_use]
    pub const fn is_changed(&self) -> bool {
        self.changed()
    }

    /// Borrow the published snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }

    /// Consume the commit into its published snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Consume the commit into its reversible patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A reversible, source-checked replacement of one complete chart graph.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Snapshot required before applying this patch.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Snapshot produced by this patch.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact no-op.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Whether this patch changes the chart graph.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.is_empty()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Return the source revision required for forward application.
    #[must_use]
    pub const fn expected_revision(&self) -> Revision {
        self.before.revision()
    }

    /// Apply this patch atomically to a package with the captured source.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<Snapshot> {
        let current = Snapshot::load(target, self.before.document_name())?;
        if !current.same_source(&self.before) {
            return Err(invalid("chart patch source is stale"));
        }
        if self.is_empty() {
            return Ok(current);
        }

        let mut candidate = target.clone();
        package::store(
            &mut candidate,
            self.after.document_name(),
            self.after.graph(),
        )?;
        let snapshot = Snapshot::load(&candidate, self.after.document_name())?;
        if !snapshot.same_source(&self.after) {
            return Err(invalid("chart patch publication changed the staged graph"));
        }
        *target = candidate;
        Ok(snapshot)
    }

    /// Apply the inverse patch atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn undo(&self, target: &mut OpcPackage) -> Result<Snapshot> {
        self.inverse().apply(target)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Source {
    document: PartState,
    owned: Vec<PartState>,
    revision: Revision,
}

impl Source {
    fn capture(package: &OpcPackage, document_name: &PackURI, graph: &Graph) -> Result<Self> {
        let document = PartState::from_part(package.get_part(document_name)?);
        let names = ownership(graph);
        let mut owned = Vec::with_capacity(names.len());
        for name in names {
            let uri = PackURI::new(&name).map_err(Error::InvalidUri)?;
            owned.push(PartState::from_part(package.get_part(&uri)?));
        }
        owned.sort_unstable_by(|left, right| left.name.as_str().cmp(right.name.as_str()));
        let revision = fingerprint(&document, &owned);
        Ok(Self {
            document,
            owned,
            revision,
        })
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct PartState {
    name: PackURI,
    content_type: String,
    data: Arc<Vec<u8>>,
    relationships: Vec<RelationshipState>,
}

impl PartState {
    fn from_part(part: &dyn Part) -> Self {
        let mut relationships = part
            .rels()
            .iter()
            .map(RelationshipState::from_relationship)
            .collect::<Vec<_>>();
        relationships.sort_unstable_by(|left, right| left.id.cmp(&right.id));
        Self {
            name: part.partname().clone(),
            content_type: part.content_type().to_owned(),
            data: part.blob_arc(),
            relationships,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct RelationshipState {
    id: String,
    relationship_type: String,
    target_ref: String,
    target_mode: TargetMode,
}

impl RelationshipState {
    fn from_relationship(relationship: &litchi_opc::Relationship) -> Self {
        Self {
            id: relationship.r_id().to_owned(),
            relationship_type: relationship.reltype().to_owned(),
            target_ref: relationship.target_ref().to_owned(),
            target_mode: relationship.target_mode(),
        }
    }
}

fn fingerprint(document: &PartState, owned: &[PartState]) -> Revision {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    feed_part(&mut hash, document);
    for part in owned {
        feed_part(&mut hash, part);
    }
    hash
}

fn feed_part(hash: &mut Revision, part: &PartState) {
    feed(hash, part.name.as_str().as_bytes());
    feed(hash, part.content_type.as_bytes());
    feed(hash, part.data.as_slice());
    for relationship in &part.relationships {
        feed(hash, relationship.id.as_bytes());
        feed(hash, relationship.relationship_type.as_bytes());
        feed(hash, relationship.target_ref.as_bytes());
        feed(
            hash,
            &[match relationship.target_mode {
                TargetMode::Internal => 0,
                TargetMode::External => 1,
            }],
        );
    }
}

fn feed(hash: &mut Revision, bytes: &[u8]) {
    for byte in bytes {
        *hash ^= u64::from(*byte);
        *hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
}
