//! Detached, source-checked edits for the validated presentation structure.
//!
//! The presentation slide relationship graph is captured as immutable source
//! context. Detached edits only the semantic custom-show and section
//! layers, while publication replaces one presentation XML blob atomically.
//! Every candidate is parsed and validated again before it becomes a commit.

use std::sync::Arc;

use litchi_opc::{OpcPackage, Part, TargetMode};

use super::codec;
use super::model::{Graph, Reference};
use crate::presentation_properties::metadata::Limits;
use crate::presentation_properties::metadata::custom_show::List as ShowList;
use crate::presentation_properties::metadata::sections::List as SectionList;
use crate::{Error, Result};

/// Stable fingerprint of the exact presentation XML source bytes.
pub type Revision = u64;

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
    pub(crate) limits: Limits,
    pub(crate) revision: Revision,
}

impl Snapshot {
    /// Load and validate the presentation structure graph.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn load(package: &OpcPackage, limits: Limits) -> Result<Self> {
        super::package::load_snapshot(package, limits)
    }

    pub(crate) fn from_wire(
        presentation_part_name: String,
        presentation_content_type: String,
        source_xml: Arc<Vec<u8>>,
        relationships: Vec<SourceRelationship>,
        graph: Graph,
        limits: Limits,
    ) -> Result<Self> {
        if source_xml.len() > limits.max_input_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::InputBytes),
                observed: source_xml.len(),
                maximum: limits.max_input_bytes(),
            });
        }
        codec::validate_detached_graph(&graph, &graph, limits)?;
        Ok(Self {
            presentation_part_name,
            presentation_content_type,
            revision: fingerprint(source_xml.as_slice()),
            source_xml,
            relationships,
            graph,
            limits,
        })
    }

    /// Borrow the complete validated presentation graph.
    #[inline]
    #[must_use]
    pub fn graph(&self) -> &Graph {
        &self.graph
    }

    /// Borrow the resolved, ordered presentation slide references.
    #[inline]
    #[must_use]
    pub fn slides(&self) -> &[Reference] {
        &self.graph.slides
    }

    /// Borrow the typed custom-show catalog.
    #[inline]
    #[must_use]
    pub fn custom_shows(&self) -> &ShowList {
        &self.graph.custom_shows
    }

    /// Borrow the typed section catalog.
    #[inline]
    #[must_use]
    pub fn sections(&self) -> &SectionList {
        &self.graph.sections
    }

    /// Return the owning `PresentationML` part name.
    #[inline]
    #[must_use]
    pub fn presentation_part_name(&self) -> &str {
        &self.presentation_part_name
    }

    /// Return the source `PresentationML` content type.
    #[inline]
    #[must_use]
    pub fn presentation_content_type(&self) -> &str {
        &self.presentation_content_type
    }

    /// Return the source fingerprint used for stale-source checks.
    #[inline]
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Borrow the exact presentation XML captured by this snapshot.
    #[inline]
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    /// Start an atomic detached edit over custom shows and sections.
    #[inline]
    #[must_use]
    pub(crate) fn edit(&self) -> Draft {
        Draft {
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
pub struct Draft {
    original: Snapshot,
    working: Graph,
}

impl Draft {
    /// Immutable source snapshot used for conflict checks and inverse patches.
    #[inline]
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.original
    }

    /// Borrow the currently staged graph.
    #[inline]
    #[must_use]
    pub fn graph(&self) -> &Graph {
        &self.working
    }

    /// Borrow the staged custom-show catalog.
    #[inline]
    #[must_use]
    pub fn custom_shows(&self) -> &ShowList {
        &self.working.custom_shows
    }

    /// Borrow the staged section catalog.
    #[inline]
    #[must_use]
    pub fn sections(&self) -> &SectionList {
        &self.working.sections
    }

    fn has_changes(&self) -> bool {
        !codec::equivalent_graph(&self.original.graph, &self.working)
    }

    /// The outcome of this edit or commit.
    #[must_use]
    pub fn diagnostics(&self) -> crate::diagnostics::Diagnostics {
        crate::diagnostics::Diagnostics::new(self.has_changes())
    }

    /// Replace the complete typed graph, retaining the source slide topology.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace(&mut self, value: Graph) -> Result<bool> {
        self.validate(&value)?;
        if codec::equivalent_graph(&self.working, &value) {
            return Ok(false);
        }
        self.working = value;
        Ok(true)
    }

    /// Apply a checked mutation to a cloned graph without partial staging.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn edit(&mut self, edit: impl FnOnce(&mut Graph) -> Result<()>) -> Result<()> {
        let mut candidate = self.working.clone();
        edit(&mut candidate)?;
        self.validate(&candidate)?;
        self.working = candidate;
        Ok(())
    }

    /// Apply a checked mutation to only the custom-show catalog.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn edit_custom_shows(
        &mut self,
        edit: impl FnOnce(&mut ShowList) -> Result<()>,
    ) -> Result<()> {
        self.edit(|graph| edit(&mut graph.custom_shows))
    }

    /// Apply a checked mutation to only the section catalog.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn edit_sections(
        &mut self,
        edit: impl FnOnce(&mut SectionList) -> Result<()>,
    ) -> Result<()> {
        self.edit(|graph| edit(&mut graph.sections))
    }

    /// Replace the complete custom-show catalog after graph validation.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_custom_shows(&mut self, value: ShowList) -> Result<bool> {
        let before = self.working.clone();
        self.edit_custom_shows(|shows| {
            *shows = value;
            Ok(())
        })?;
        Ok(!codec::equivalent_graph(&before, &self.working))
    }

    /// Replace the complete section catalog after graph validation.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_sections(&mut self, value: SectionList) -> Result<bool> {
        let before = self.working.clone();
        self.edit_sections(|sections| {
            *sections = value;
            Ok(())
        })?;
        Ok(!codec::equivalent_graph(&before, &self.working))
    }

    /// Validate and consume the edit into a source-checked commit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn commit(self) -> Result<Prepared> {
        if !self.has_changes() {
            let patch = Change::new(self.original.clone(), self.original.clone());
            return Ok(Prepared {
                snapshot: self.original,
                patch,
                changed: false,
            });
        }

        let updated = codec::rewrite(
            self.original.source_xml.as_slice(),
            &self.original.graph,
            &self.working,
            self.original.limits,
        )?;
        let graph = codec::parse_detached(&updated, &self.working.slides, self.original.limits)?;
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
            self.original.limits,
        )?;
        let patch = Change::new(self.original, snapshot.clone());
        Ok(Prepared {
            snapshot,
            patch,
            changed: true,
        })
    }

    fn validate(&self, value: &Graph) -> Result<()> {
        codec::validate_detached_graph(&self.original.graph, value, self.original.limits)
    }
}

/// A successful presentation-structure edit and its reversible patch.
#[derive(Clone, Debug)]
pub(crate) struct Prepared {
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the module's source-level tests")
    )]
    snapshot: Snapshot,
    patch: Change,
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the module's source-level tests")
    )]
    changed: bool,
}

impl Prepared {
    #[cfg(test)]
    fn has_changes(&self) -> bool {
        self.changed
    }

    /// The outcome of this edit or commit.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn diagnostics(&self) -> crate::diagnostics::Diagnostics {
        crate::diagnostics::Diagnostics::new(self.has_changes())
    }

    /// Borrow the projected post-edit snapshot.
    #[inline]
    #[must_use]
    #[cfg(test)]
    pub(crate) fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible source-checked patch.
    #[inline]
    #[must_use]
    pub(crate) fn patch(&self) -> &Change {
        &self.patch
    }

    /// Consume the commit into its patch.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn into_patch(self) -> Change {
        self.patch
    }
}

/// A reversible source-checked replacement of one presentation XML blob.
#[derive(Clone, Debug)]
pub(crate) struct Change {
    before: Snapshot,
    after: Snapshot,
}

impl Change {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source context required before publication.
    #[inline]
    #[must_use]
    pub(crate) fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Source context produced by publication.
    #[inline]
    #[must_use]
    pub(crate) fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact source no-op.
    #[inline]
    #[must_use]
    pub(crate) fn is_noop(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the exact inverse patch without reinterpreting XML.
    #[inline]
    #[must_use]
    #[cfg(test)]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply this patch atomically after checking the source and relationship topology.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
    /// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(test)]
    pub(crate) fn apply(&self, target: &mut OpcPackage) -> Result<Snapshot> {
        super::package::apply_patch(target, self)
    }
}

fn fingerprint(bytes: &[u8]) -> Revision {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100_0000_01b3);
    }
    hash
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
