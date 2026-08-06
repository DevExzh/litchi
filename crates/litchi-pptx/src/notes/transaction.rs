//! Source-checked semantic transactions for an existing notes graph.

use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI, Part, TargetMode};

use super::codec::{self, validate_resource_xml};
use super::model::{Graph, Link};
use super::validation;
use super::{MAX_MASTER_XML, MAX_NOTES_XML, MAX_THEME_XML, invalid};
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
}

impl Snapshot {
    /// Load a source-bound snapshot when the presentation has notes.
    pub fn load(package: &OpcPackage, presentation_name: &PackURI) -> Result<Option<Self>> {
        super::package::load_snapshot(package, presentation_name)
    }

    /// Alias emphasizing that the returned value is tied to source bytes.
    pub fn read(package: &OpcPackage, presentation_name: &PackURI) -> Result<Option<Self>> {
        Self::load(package, presentation_name)
    }

    pub(crate) fn from_parts(
        presentation_part_name: PackURI,
        presentation: PartState,
        mut parts: Vec<PartState>,
        graph: Graph,
    ) -> Result<Self> {
        parts.sort_unstable_by(|left, right| left.name.as_str().cmp(right.name.as_str()));
        validation::validate_parts(&parts)?;
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
        })
    }

    /// The physical PresentationML part owning this graph.
    pub fn presentation_part_name(&self) -> &PackURI {
        &self.presentation_part_name
    }

    /// Exact source PresentationML bytes captured by this snapshot.
    pub fn source_xml(&self) -> &[u8] {
        self.presentation.data.as_slice()
    }

    /// Borrow the validated semantic graph.
    pub fn graph(&self) -> &Graph {
        &self.graph
    }

    /// Borrow the shared notes master.
    pub fn master(&self) -> &super::model::Master {
        self.graph.master()
    }

    /// Borrow notes slides in PresentationML order.
    pub fn slides(&self) -> &[super::model::Slide] {
        self.graph.slides()
    }

    /// Compact source revision used for optimistic stale checks.
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Start a detached atomic edit over this snapshot.
    pub fn edit(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            working: self.graph.clone(),
        }
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.presentation_part_name == other.presentation_part_name
            && self.parts == other.parts
            && self.graph == other.graph
            && self.revision == other.revision
    }
}

/// Failure-atomic editor for inert notes text, master XML, and theme XML.
#[derive(Clone, Debug)]
pub struct Transaction {
    source: Snapshot,
    working: Graph,
}

impl Transaction {
    /// Borrow the immutable source snapshot used by this edit.
    pub fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Borrow the currently staged semantic graph.
    pub fn graph(&self) -> &Graph {
        &self.working
    }

    /// Whether any staged resource bytes or graph metadata differ from source.
    pub fn is_changed(&self) -> bool {
        self.working != self.source.graph
    }

    /// Replace one notes slide's inert text projection while preserving XML.
    pub fn set_text(&mut self, index: usize, text: impl AsRef<str>) -> Result<bool> {
        let text = text.as_ref();
        let slide = self
            .working
            .slides()
            .get(index)
            .ok_or_else(|| invalid("notes slide index is out of bounds"))?;
        let current = slide.text()?;
        if current.as_deref() == (!text.is_empty()).then_some(text) {
            return Ok(false);
        }
        let xml = codec::rewrite_text(slide.xml(), text)?;
        self.working.slides_mut()[index].replace_xml(xml);
        Ok(true)
    }

    /// Alias emphasizing that text is owned by one notes slide.
    pub fn set_slide_text(&mut self, index: usize, text: impl AsRef<str>) -> Result<bool> {
        self.set_text(index, text)
    }

    /// Replace one notes slide's XML after bounded conformance validation.
    pub fn replace_slide_xml(&mut self, index: usize, xml: Vec<u8>) -> Result<bool> {
        let conformance = self.working.conformance();
        validate_resource_xml(&xml, MAX_NOTES_XML, conformance, "notes", "notes slide")?;
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
    pub fn replace_master_xml(&mut self, xml: Vec<u8>) -> Result<bool> {
        let conformance = self.working.conformance();
        validate_resource_xml(
            &xml,
            MAX_MASTER_XML,
            conformance,
            "notesMaster",
            "notes master",
        )?;
        if self.working.master().xml() == xml {
            return Ok(false);
        }
        self.working.master_mut().replace_xml(xml);
        Ok(true)
    }

    /// Replace the inert notes-master theme XML while retaining all edges.
    pub fn replace_theme_xml(&mut self, xml: Vec<u8>) -> Result<bool> {
        let conformance = self.working.conformance();
        validate_resource_xml(&xml, MAX_THEME_XML, conformance, "theme", "notes theme")?;
        if self.working.master().theme().xml() == xml {
            return Ok(false);
        }
        self.working.master_mut().theme_mut().replace_xml(xml);
        Ok(true)
    }

    /// Validate and consume this edit into a reversible source patch.
    pub fn commit(self) -> Result<Commit> {
        super::package::validate_graph(&self.working)?;
        let changed = self.is_changed();
        if !changed {
            let patch = Patch {
                before: self.source.clone(),
                after: self.source.clone(),
            };
            return Ok(Commit {
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
        )?;
        let patch = Patch {
            before: self.source,
            after: snapshot.clone(),
        };
        Ok(Commit { snapshot, patch })
    }

    /// Discard edits and recover the exact source snapshot.
    pub fn rollback(self) -> Snapshot {
        self.source
    }
}

/// A successful detached edit and its reversible package patch.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Snapshot expected after publication.
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Source-checked patch represented by this commit.
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Whether publication changes any captured graph bytes.
    pub fn is_changed(&self) -> bool {
        !self.patch.is_empty()
    }

    /// Alias for callers that prefer a verb-style result.
    pub fn changed(&self) -> bool {
        self.is_changed()
    }

    /// Consume the commit into its target snapshot and patch.
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }

    /// Consume the commit into its patch.
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A reversible source-checked replacement of one existing notes graph.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    /// Snapshot required before forward application.
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Snapshot expected after forward application.
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact no-op.
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Whether this patch changes the source graph.
    pub fn is_changed(&self) -> bool {
        !self.is_empty()
    }

    /// Return the exact inverse patch.
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Source fingerprint required for publication.
    pub const fn expected_revision(&self) -> Revision {
        self.before.revision
    }

    /// Publish this patch atomically to an OPC package.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        super::package::apply_patch(package, self)
    }

    /// Apply the inverse patch to restore the exact source graph.
    pub fn undo(&self, package: &mut OpcPackage) -> Result<Snapshot> {
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
