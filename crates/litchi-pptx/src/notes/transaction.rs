//! Source-checked semantic transactions for an existing notes graph.

use std::sync::Arc;

#[cfg(test)]
use litchi_opc::OpcPackage;
use litchi_opc::{PackURI, Part, TargetMode};

use super::codec::{self, validate_resource_xml_with_limits};
use super::model::{Graph, Link};
use super::validation;
use super::{Limits, MAX_MASTER_XML, MAX_NOTES_XML, MAX_THEME_XML, invalid};
use crate::Result;

/// Stable fingerprint of all captured notes graph part bytes and edges.
pub type Revision = u64;

/// Exact bytes and relationships for one part owned by the notes graph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct PartState {
    pub(crate) name: PackURI,
    pub(crate) content_type: String,
    pub(crate) data: Arc<Vec<u8>>,
    pub(crate) relationships: Vec<Link>,
}

impl PartState {
    pub(crate) fn from_part(part: &dyn Part) -> Self {
        let mut relationships: Vec<_> = part
            .rels()
            .iter()
            .map(|relationship| {
                Link::new(
                    relationship.r_id(),
                    relationship.reltype(),
                    relationship.target_ref(),
                    relationship.target_mode(),
                )
            })
            .collect();
        relationships.sort_unstable_by(|left, right| left.id.cmp(&right.id));
        Self {
            name: part.partname().clone(),
            content_type: part.content_type().to_owned(),
            data: part.blob_arc(),
            relationships,
        }
    }
}

/// An immutable semantic snapshot of one validated notes graph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    pub(crate) presentation_part_name: PackURI,
    pub(crate) presentation: PartState,
    pub(crate) parts: Vec<PartState>,
    pub(crate) graph: Graph,
    pub(crate) revision: Revision,
    pub(crate) limits: Limits,
}

impl Snapshot {
    /// Load a source-bound snapshot when the presentation has notes.
    #[cfg(test)]
    pub(crate) fn load(package: &OpcPackage, presentation_name: &PackURI) -> Result<Option<Self>> {
        super::package::load_snapshot(package, presentation_name)
    }

    pub(crate) fn from_parts(
        presentation_part_name: PackURI,
        presentation: PartState,
        mut parts: Vec<PartState>,
        graph: Graph,
        limits: Limits,
    ) -> Result<Self> {
        parts.sort_unstable_by(|left, right| left.name.as_str().cmp(right.name.as_str()));
        validation::validate_parts_with_limits(&parts, limits)?;
        if presentation.name != presentation_part_name {
            return Err(invalid(
                "notes presentation source identity is inconsistent",
            ));
        }
        let revision = fingerprint(&parts);
        Ok(Self {
            presentation_part_name,
            presentation,
            parts,
            graph,
            revision,
            limits,
        })
    }

    /// The physical `PresentationML` part owning this graph.
    #[must_use]
    pub fn presentation_part_name(&self) -> &PackURI {
        &self.presentation_part_name
    }

    /// Exact source `PresentationML` bytes captured by this snapshot.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.presentation.data.as_slice()
    }

    /// Borrow the validated semantic graph.
    #[must_use]
    pub fn graph(&self) -> &Graph {
        &self.graph
    }

    /// Borrow the shared notes master.
    #[must_use]
    pub fn master(&self) -> &super::model::Master {
        self.graph.master()
    }

    /// Borrow notes slides in `PresentationML` order.
    #[must_use]
    pub fn slides(&self) -> &[super::model::Slide] {
        self.graph.slides()
    }

    /// Compact source revision used for optimistic stale checks.
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Start a detached atomic edit over this snapshot.
    #[must_use]
    pub(crate) fn edit(&self) -> Draft {
        Draft {
            source: self.clone(),
            working: self.graph.clone(),
        }
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.presentation_part_name == other.presentation_part_name
            && self.parts == other.parts
            && self.graph == other.graph
            && self.revision == other.revision
            && self.limits == other.limits
    }
}

/// Failure-atomic editor for inert notes text, master XML, and theme XML.
#[derive(Clone, Debug)]
pub(crate) struct Draft {
    source: Snapshot,
    working: Graph,
}

impl Draft {
    fn is_changed(&self) -> bool {
        self.working != self.source.graph
    }

    /// Replace the complete detached graph without changing physical owners.
    ///
    /// # Errors
    ///
    /// Returns an error when the replacement is malformed or retargets any
    /// presentation, master, theme, slide, or notes relationship identity.
    pub(crate) fn replace(&mut self, graph: Graph) -> Result<bool> {
        super::package::validate_graph_with_limits(&graph, self.source.limits)?;
        if !same_ownership(&self.source.graph, &graph) {
            return Err(invalid(
                "notes transaction cannot retarget or orphan physical owners",
            ));
        }
        if self.working == graph {
            return Ok(false);
        }
        self.working = graph;
        Ok(true)
    }

    /// Replace one notes slide's inert text projection while preserving XML.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn set_text(&mut self, index: usize, text: impl AsRef<str>) -> Result<bool> {
        let text = text.as_ref();
        let slide = self
            .working
            .slides()
            .get(index)
            .ok_or_else(|| invalid("notes slide index is out of bounds"))?;
        let current = slide.text(self.source.limits)?;
        if current.as_deref() == (!text.is_empty()).then_some(text) {
            return Ok(false);
        }
        let xml = codec::rewrite_text(slide.xml(), text, self.source.limits)?;
        self.working.slides_mut()[index].replace_xml(xml);
        Ok(true)
    }

    /// Replace one notes slide's XML after bounded conformance validation.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn replace_slide_xml(&mut self, index: usize, xml: Vec<u8>) -> Result<bool> {
        let conformance = self.working.conformance();
        validate_resource_xml_with_limits(
            &xml,
            MAX_NOTES_XML,
            conformance,
            "notes",
            "notes slide",
            self.source.limits,
        )?;
        let slide = self
            .working
            .slides()
            .get(index)
            .ok_or_else(|| invalid("notes slide index is out of bounds"))?;
        if slide.xml() == xml {
            return Ok(false);
        }
        self.working.slides_mut()[index].replace_xml(xml);
        Ok(true)
    }

    /// Replace the inert notes-master XML while retaining its relationships.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn replace_master_xml(&mut self, xml: Vec<u8>) -> Result<bool> {
        let conformance = self.working.conformance();
        validate_resource_xml_with_limits(
            &xml,
            MAX_MASTER_XML,
            conformance,
            "notesMaster",
            "notes master",
            self.source.limits,
        )?;
        if self.working.master().xml() == xml {
            return Ok(false);
        }
        self.working.master_mut().replace_xml(xml);
        Ok(true)
    }

    /// Replace the inert notes-master theme XML while retaining all edges.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn replace_theme_xml(&mut self, xml: Vec<u8>) -> Result<bool> {
        let conformance = self.working.conformance();
        validate_resource_xml_with_limits(
            &xml,
            MAX_THEME_XML,
            conformance,
            "theme",
            "notes theme",
            self.source.limits,
        )?;
        if self.working.master().theme().xml() == xml {
            return Ok(false);
        }
        self.working.master_mut().theme_mut().replace_xml(xml);
        Ok(true)
    }

    /// Validate and consume this edit into a reversible source patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn commit(self) -> Result<Prepared> {
        super::package::validate_graph_with_limits(&self.working, self.source.limits)?;
        let changed = self.is_changed();
        if !changed {
            let patch = Change {
                before: self.source.clone(),
                after: self.source.clone(),
            };
            return Ok(Prepared {
                snapshot: self.source,
                patch,
            });
        }
        let parts = parts_after(&self.source, &self.working)?;
        let snapshot = Snapshot::from_parts(
            self.source.presentation_part_name.clone(),
            self.source.presentation.clone(),
            parts,
            self.working,
            self.source.limits,
        )?;
        let patch = Change {
            before: self.source,
            after: snapshot.clone(),
        };
        Ok(Prepared { snapshot, patch })
    }
}

/// A successful detached edit and its reversible package patch.
#[derive(Clone, Debug)]
pub(crate) struct Prepared {
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the module's source-level tests")
    )]
    snapshot: Snapshot,
    patch: Change,
}

impl Prepared {
    /// Snapshot expected after publication.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Source-checked patch represented by this commit.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn patch(&self) -> &Change {
        &self.patch
    }

    fn is_changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// The outcome of this edit or commit.
    #[must_use]
    pub(crate) fn diagnostics(&self) -> crate::diagnostics::Diagnostics {
        crate::diagnostics::Diagnostics::new(self.is_changed())
    }

    /// Consume the commit into its patch.
    #[must_use]
    pub(crate) fn into_patch(self) -> Change {
        self.patch
    }
}

/// A reversible source-checked replacement of one existing notes graph.
#[derive(Clone, Debug)]
pub(crate) struct Change {
    before: Snapshot,
    after: Snapshot,
}

impl Change {
    /// Snapshot required before forward application.
    #[must_use]
    pub(crate) fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Snapshot expected after forward application.
    #[must_use]
    pub(crate) fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact no-op.
    #[must_use]
    pub(crate) fn is_noop(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the exact inverse patch.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Publish this patch atomically to an OPC package.
    #[cfg(test)]
    pub(crate) fn apply(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        super::package::apply_patch(package, self)
    }

    /// Apply the inverse patch to restore the exact source graph.
    #[cfg(test)]
    pub(crate) fn undo(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        self.inverse().apply(package)
    }
}

fn parts_after(source: &Snapshot, graph: &Graph) -> Result<Vec<PartState>> {
    let mut parts = source.parts.clone();
    replace_part(
        &mut parts,
        &graph.master().part_name,
        graph.master().xml(),
        graph.master().relationships(),
    )?;
    replace_part(
        &mut parts,
        &graph.master().theme().part_name,
        graph.master().theme().xml(),
        graph.master().theme().relationships(),
    )?;
    for slide in graph.slides() {
        replace_part(
            &mut parts,
            &slide.part_name,
            slide.xml(),
            slide.relationships(),
        )?;
    }
    Ok(parts)
}

fn replace_part(parts: &mut [PartState], name: &str, data: &[u8], links: &[Link]) -> Result<()> {
    let part = parts
        .iter_mut()
        .find(|part| part.name.as_str() == name)
        .ok_or_else(|| invalid("notes transaction part identity changed"))?;
    part.data = Arc::new(data.to_vec());
    part.relationships = links.to_vec();
    Ok(())
}

fn same_ownership(left: &Graph, right: &Graph) -> bool {
    left.conformance == right.conformance
        && left.master.presentation_relationship_id == right.master.presentation_relationship_id
        && left.master.part_name == right.master.part_name
        && left.master.theme.relationship_id == right.master.theme.relationship_id
        && left.master.theme.part_name == right.master.theme.part_name
        && left.slides.len() == right.slides.len()
        && left.slides.iter().zip(&right.slides).all(|(left, right)| {
            left.slide_part_name == right.slide_part_name
                && left.slide_relationship_id == right.slide_relationship_id
                && left.part_name == right.part_name
                && left.backlink_relationship_id == right.backlink_relationship_id
                && left.notes_master_relationship_id == right.notes_master_relationship_id
        })
}

fn fingerprint(parts: &[PartState]) -> Revision {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    for part in parts {
        feed(&mut hash, part.name.as_str().as_bytes());
        feed(&mut hash, part.content_type.as_bytes());
        feed(&mut hash, part.data.as_slice());
        for link in &part.relationships {
            feed(&mut hash, link.id.as_bytes());
            feed(&mut hash, link.relationship_type.as_bytes());
            feed(&mut hash, link.target_ref.as_bytes());
            feed(
                &mut hash,
                &[match link.target_mode {
                    TargetMode::Internal => 0,
                    TargetMode::External => 1,
                }],
            );
        }
    }
    hash
}

fn feed(hash: &mut u64, bytes: &[u8]) {
    for byte in bytes {
        *hash ^= u64::from(*byte);
        *hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
}
