//! Detached, source-checked edits for the validated presentation structure.
//!
//! The presentation slide relationship graph is captured as immutable source
//! context. Transactions edit only the semantic custom-show and section
//! layers, while publication replaces one presentation XML blob atomically.
//! Every candidate is parsed and validated again before it becomes a commit.

use std::sync::Arc;

use litchi_opc::{OpcPackage, Part, TargetMode};

use super::codec;
use super::model::{Graph, Reference};
use crate::presentation_properties::metadata::custom_show::List as ShowList;
use crate::presentation_properties::metadata::sections::List as SectionList;
use crate::{Error, Result};

/// Stable fingerprint of the exact presentation XML source bytes.
pub type Revision = u64;

const MAX_SOURCE_BYTES: usize = 8 * 1024 * 1024;

/// One relationship entry captured from the owning presentation part.
///
/// This is deliberately package-private: it is source context for conflict
/// detection, not a second public relationship model.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourceRelationship {
    pub(crate) id: String,
    pub(crate) relationship_type: String,
    pub(crate) target: String,
    pub(crate) target_mode: TargetMode,
}

impl SourceRelationship {
    fn from_opc(value: &litchi_opc::Relationship) -> Self {
        Self {
            id: value.r_id().to_owned(),
            relationship_type: value.reltype().to_owned(),
            target: value.target_ref().to_owned(),
            target_mode: value.target_mode(),
        }
    }
}

/// An immutable semantic graph bound to one exact presentation source.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    pub(crate) presentation_part_name: String,
    pub(crate) presentation_content_type: String,
    pub(crate) source_xml: Arc<Vec<u8>>,
    pub(crate) relationships: Vec<SourceRelationship>,
    pub(crate) graph: Graph,
    pub(crate) revision: Revision,
}

impl Snapshot {
    /// Load and validate the presentation structure graph.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        super::package::load_snapshot(package)
    }

    /// Alias for [`Self::load`] emphasizing the source-bound result.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        Self::load(package)
    }

    pub(crate) fn from_wire(
        presentation_part_name: String,
        presentation_content_type: String,
        source_xml: Arc<Vec<u8>>,
        relationships: Vec<SourceRelationship>,
        graph: Graph,
    ) -> Result<Self> {
        if source_xml.len() > MAX_SOURCE_BYTES {
            return Err(limit("presentation structure source bytes"));
        }
        codec::validate_detached_graph(&graph, &graph)?;
        Ok(Self {
            presentation_part_name,
            presentation_content_type,
            revision: fingerprint(source_xml.as_slice()),
            source_xml,
            relationships,
            graph,
        })
    }

    /// Borrow the complete validated presentation graph.
    #[inline]
    pub fn graph(&self) -> &Graph {
        &self.graph
    }

    /// Contextual alias for [`Self::graph`].
    #[inline]
    pub fn structure(&self) -> &Graph {
        self.graph()
    }

    /// Borrow the resolved, ordered presentation slide references.
    #[inline]
    pub fn slides(&self) -> &[Reference] {
        &self.graph.slides
    }

    /// Borrow the typed custom-show catalog.
    #[inline]
    pub fn custom_shows(&self) -> &ShowList {
        &self.graph.custom_shows
    }

    /// Borrow the typed section catalog.
    #[inline]
    pub fn sections(&self) -> &SectionList {
        &self.graph.sections
    }

    /// Return the owning PresentationML part name.
    #[inline]
    pub fn presentation_part_name(&self) -> &str {
        &self.presentation_part_name
    }

    /// Return the source PresentationML content type.
    #[inline]
    pub fn presentation_content_type(&self) -> &str {
        &self.presentation_content_type
    }

    /// Return the source fingerprint used for stale-source checks.
    #[inline]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Borrow the exact presentation XML captured by this snapshot.
    #[inline]
    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    /// Start an atomic detached edit over custom shows and sections.
    #[inline]
    pub fn edit(&self) -> Transaction {
        Transaction {
            original: self.clone(),
            working: self.graph.clone(),
        }
    }

    pub(crate) fn capture_relationships(part: &dyn Part) -> Vec<SourceRelationship> {
        let mut relationships = part
            .rels()
            .iter()
            .map(SourceRelationship::from_opc)
            .collect::<Vec<_>>();
        relationships.sort_unstable_by(|left, right| left.id.cmp(&right.id));
        relationships
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.presentation_part_name == other.presentation_part_name
            && self.presentation_content_type == other.presentation_content_type
            && self.source_xml.as_slice() == other.source_xml.as_slice()
            && self.relationships == other.relationships
            && self.revision == other.revision
    }

    pub(crate) fn source_arc(&self) -> &Arc<Vec<u8>> {
        &self.source_xml
    }
}

/// A detached atomic edit over custom shows and sections.
#[derive(Clone, Debug)]
pub struct Transaction {
    original: Snapshot,
    working: Graph,
}

impl Transaction {
    /// Immutable source snapshot used for conflict checks and inverse patches.
    #[inline]
    pub fn before(&self) -> &Snapshot {
        &self.original
    }

    /// Borrow the currently staged graph.
    #[inline]
    pub fn graph(&self) -> &Graph {
        &self.working
    }

    /// Contextual alias for [`Self::graph`].
    #[inline]
    pub fn snapshot(&self) -> &Graph {
        self.graph()
    }

    /// Borrow the staged custom-show catalog.
    #[inline]
    pub fn custom_shows(&self) -> &ShowList {
        &self.working.custom_shows
    }

    /// Borrow the staged section catalog.
    #[inline]
    pub fn sections(&self) -> &SectionList {
        &self.working.sections
    }

    /// Whether the staged semantic graph differs from the captured source.
    #[inline]
    pub fn is_changed(&self) -> bool {
        !codec::equivalent_graph(&self.original.graph, &self.working)
    }

    /// Replace the complete typed graph, retaining the source slide topology.
    pub fn replace(&mut self, value: Graph) -> Result<bool> {
        self.validate(&value)?;
        if codec::equivalent_graph(&self.working, &value) {
            return Ok(false);
        }
        self.working = value;
        Ok(true)
    }

    /// Apply a checked mutation to a cloned graph without partial staging.
    pub fn edit(&mut self, edit: impl FnOnce(&mut Graph) -> Result<()>) -> Result<()> {
        let mut candidate = self.working.clone();
        edit(&mut candidate)?;
        self.validate(&candidate)?;
        self.working = candidate;
        Ok(())
    }

    /// Apply a checked mutation to only the custom-show catalog.
    pub fn edit_custom_shows(
        &mut self,
        edit: impl FnOnce(&mut ShowList) -> Result<()>,
    ) -> Result<()> {
        self.edit(|graph| edit(&mut graph.custom_shows))
    }

    /// Apply a checked mutation to only the section catalog.
    pub fn edit_sections(
        &mut self,
        edit: impl FnOnce(&mut SectionList) -> Result<()>,
    ) -> Result<()> {
        self.edit(|graph| edit(&mut graph.sections))
    }

    /// Replace the complete custom-show catalog after graph validation.
    pub fn set_custom_shows(&mut self, value: ShowList) -> Result<bool> {
        let before = self.working.clone();
        self.edit_custom_shows(|shows| {
            *shows = value;
            Ok(())
        })?;
        Ok(!codec::equivalent_graph(&before, &self.working))
    }

    /// Replace the complete section catalog after graph validation.
    pub fn set_sections(&mut self, value: SectionList) -> Result<bool> {
        let before = self.working.clone();
        self.edit_sections(|sections| {
            *sections = value;
            Ok(())
        })?;
        Ok(!codec::equivalent_graph(&before, &self.working))
    }

    /// Validate and consume the edit into a source-checked commit.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.original.clone(), self.original.clone());
            return Ok(Commit {
                snapshot: self.original,
                patch,
                changed: false,
            });
        }

        let updated = codec::rewrite(
            self.original.source_xml.as_slice(),
            &self.original.graph,
            &self.working,
        )?;
        let graph = codec::parse_detached(&updated, &self.working.slides)?;
        if !codec::equivalent_graph(&graph, &self.working) {
            return Err(invalid(
                "published presentation structure changed the staged graph",
            ));
        }
        let snapshot = Snapshot::from_wire(
            self.original.presentation_part_name.clone(),
            self.original.presentation_content_type.clone(),
            Arc::new(updated),
            self.original.relationships.clone(),
            graph,
        )?;
        let patch = Patch::new(self.original, snapshot.clone());
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }

    fn validate(&self, value: &Graph) -> Result<()> {
        codec::validate_detached_graph(&self.original.graph, value)
    }
}

/// A successful presentation-structure edit and its reversible patch.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether publication changes the exact presentation XML bytes.
    #[inline]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Alias for [`Self::changed`].
    #[inline]
    pub const fn is_changed(&self) -> bool {
        self.changed
    }

    /// Borrow the projected post-edit snapshot.
    #[inline]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible source-checked patch.
    #[inline]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its snapshot and patch.
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }

    /// Consume the commit into its patch.
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A reversible source-checked replacement of one presentation XML blob.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source context required before publication.
    #[inline]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Source context produced by publication.
    #[inline]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact source no-op.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Alias for [`Self::is_empty`].
    #[inline]
    pub fn is_noop(&self) -> bool {
        self.is_empty()
    }

    /// Whether this patch changes the presentation source.
    #[inline]
    pub fn is_changed(&self) -> bool {
        !self.is_empty()
    }

    /// Return the exact inverse patch without reinterpreting XML.
    #[inline]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Return the source fingerprint required for publication.
    #[inline]
    pub const fn expected_revision(&self) -> Revision {
        self.before.revision
    }

    /// Apply this patch atomically after checking the source and relationship topology.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<Snapshot> {
        super::package::apply_patch(target, self)
    }
}

fn fingerprint(bytes: &[u8]) -> Revision {
    let mut hash = 0xcbf29ce484222325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100000001b3);
    }
    hash
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(what: &str) -> Error {
    Error::Invalid(format!("{what} exceeds the supported safety limit"))
}
