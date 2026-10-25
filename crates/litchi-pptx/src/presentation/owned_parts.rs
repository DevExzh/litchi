//! Orphan proof for the Parts a removed slide owns exclusively (record 0303).
//!
//! A slide reaches its notes slide, pictures, diagrams, VML drawings,
//! comments and embedded objects through its own relationships. Removing the
//! slide may remove such a Part only when nothing that survives the removal
//! still targets it. The proof works on the typed relationship graph of the
//! whole package:
//!
//! 1. collect every Part reachable from the slide through internal edges,
//!    except Parts of the presentation structure (slides, layouts, masters,
//!    notes and handout masters, themes, the presentation itself), which a
//!    slide never owns;
//! 2. keep the greatest subset in which every incoming edge comes from the
//!    slide or from another kept Part. The package root's own relationships
//!    count as an outside owner.
//!
//! Every kept Part is then an orphan once the slide leaves, and every Part
//! the slide reaches but does not keep (a shared layout, a notes master, a
//! picture another slide also shows) stays. Edges from the kept Parts back to
//! the slide (a notes slide names its slide) are inside the removed set.

use std::collections::{HashMap, HashSet};

use litchi_opc::constants::content_type as ct;
use litchi_opc::{PackURI, PreparedTopology, Relationships, SourceBackedPackage};

use crate::{Error, Result, SlideRemovalRefusal};

/// Content types a slide never owns, so the proof never removes them.
const STRUCTURAL: &[&str] = &[
    ct::PML_SLIDE,
    ct::PML_SLIDE_LAYOUT,
    ct::PML_SLIDE_MASTER,
    ct::PML_NOTES_MASTER,
    ct::PML_HANDOUT_MASTER,
    ct::PML_PRESENTATION_MAIN,
    ct::PML_SLIDESHOW_MAIN,
    ct::PML_TEMPLATE_MAIN,
    ct::PML_PRES_MACRO_MAIN,
    ct::PML_SLIDESHOW_MACRO_MAIN,
    ct::PML_TEMPLATE_MACRO_MAIN,
    ct::OFC_THEME,
];

/// One internal edge, resolved against the package catalog.
struct Edge {
    /// The target Part, or `None` when no Part has that name (a dangling
    /// edge, which owns nothing).
    target: Option<usize>,
    /// The target carries a query or fragment.
    ambiguous: bool,
    /// The stored reference is the owner-relative spelling of its target.
    canonical: bool,
}

struct Node {
    name: PackURI,
    content_type: String,
    edges: Vec<Edge>,
    /// Owners of incoming internal edges; `None` is the package root.
    incoming: Vec<Option<usize>>,
}

/// Whether a slide with these relationships can own a Part: it has an
/// internal edge besides its layout and slide-jump edges. A slide without
/// one owns nothing, so callers skip the package-wide graph.
pub(crate) fn may_own(relationships: &Relationships) -> bool {
    relationships.iter().any(|relationship| {
        !relationship.is_external()
            && !crate::parts::is_relationship_type(
                relationship.reltype(),
                litchi_opc::constants::relationship_type::SLIDE_LAYOUT,
                "slideLayout",
            )
            && !crate::parts::is_relationship_type(
                relationship.reltype(),
                litchi_opc::constants::relationship_type::SLIDE,
                "slide",
            )
    })
}

/// The typed relationship graph of one package or prepared candidate.
pub(crate) struct Graph {
    nodes: Vec<Node>,
    index: HashMap<String, usize>,
}

/// The Parts a slide owns exclusively, in folded-name order.
pub(crate) struct Owned {
    parts: Vec<PackURI>,
    folded: HashSet<String>,
}

impl Owned {
    /// No owned Part.
    pub(crate) fn none() -> Self {
        Self {
            parts: Vec::new(),
            folded: HashSet::new(),
        }
    }

    /// The owned Parts, sorted by folded name.
    pub(crate) fn parts(&self) -> &[PackURI] {
        &self.parts
    }

    /// Whether `owner` is one of the owned Parts.
    pub(crate) fn contains(&self, owner: &PackURI) -> bool {
        self.folded.contains(&fold(owner.as_str()))
    }
}

impl Graph {
    /// Capture the graph of an admitted source package.
    pub(crate) fn from_source(package: &SourceBackedPackage, max_edges: usize) -> Result<Self> {
        let mut builder = Builder::new(max_edges);
        builder.root(package.rels())?;
        for part in package.iter_parts() {
            package.check_execution()?;
            builder.part(part.partname(), part.content_type(), part.rels())?;
        }
        Ok(builder.finish())
    }

    /// Capture the graph of an eager OPC package.
    pub(crate) fn from_opc(package: &litchi_opc::OpcPackage, max_edges: usize) -> Result<Self> {
        let mut builder = Builder::new(max_edges);
        builder.root(package.rels())?;
        for part in package.iter_parts() {
            builder.part(part.partname(), part.content_type(), part.rels())?;
        }
        Ok(builder.finish())
    }

    /// Capture the graph of a prepared topology candidate.
    pub(crate) fn from_prepared(prepared: &PreparedTopology, max_edges: usize) -> Result<Self> {
        let mut builder = Builder::new(max_edges);
        builder.root(prepared.rels())?;
        for part in prepared.iter_parts()? {
            prepared.check_execution()?;
            builder.part(part.partname(), part.content_type(), part.rels())?;
        }
        Ok(builder.finish())
    }

    /// The Parts `slide` owns exclusively.
    ///
    /// # Errors
    ///
    /// Refuses an owned-candidate edge with a query or fragment: removal
    /// cannot prove what such an edge names.
    pub(crate) fn owned_by(&self, slide: &PackURI) -> Result<Owned> {
        self.owned_by_all(std::slice::from_ref(slide))
    }

    /// The Parts the slides in `slides` own exclusively when they leave
    /// together: a picture two removed slides share is owned by both.
    pub(crate) fn owned_by_all(&self, slides: &[PackURI]) -> Result<Owned> {
        let mut start = vec![false; self.nodes.len()];
        let mut stack = Vec::new();
        for slide in slides {
            let Some(&node) = self.index.get(&fold(slide.as_str())) else {
                return Err(Error::PatchConflict);
            };
            start[node] = true;
            stack.push(node);
        }
        // Reachable candidates.
        let mut kept = vec![false; self.nodes.len()];
        let mut visited = start.clone();
        while let Some(current) = stack.pop() {
            for edge in &self.nodes[current].edges {
                // An edge to a missing Part (a producer's dangling diagram
                // drawing) names nothing that removal could delete.
                let Some(target) = edge.target else {
                    continue;
                };
                if start[target] || self.is_structural(target) {
                    continue;
                }
                if edge.ambiguous {
                    return refusal("a relationship of the removed slide has a query or fragment");
                }
                if !visited[target] {
                    visited[target] = true;
                    kept[target] = true;
                    stack.push(target);
                }
            }
        }
        // Greatest subset whose incoming edges all come from the slides or
        // from the subset itself.
        loop {
            let mut changed = false;
            for node in 0..self.nodes.len() {
                if !kept[node] {
                    continue;
                }
                let outside = self.nodes[node].incoming.iter().any(|owner| match owner {
                    None => true,
                    Some(owner) => !start[*owner] && !kept[*owner],
                });
                if outside {
                    kept[node] = false;
                    changed = true;
                }
            }
            if !changed {
                break;
            }
        }
        let mut parts: Vec<PackURI> = kept
            .iter()
            .enumerate()
            .filter(|(_, kept)| **kept)
            .map(|(node, _)| self.nodes[node].name.clone())
            .collect();
        parts.sort_unstable_by_key(|name| fold(name.as_str()));
        let folded = parts.iter().map(|name| fold(name.as_str())).collect();
        Ok(Owned { parts, folded })
    }

    /// Refuse a removal that a durable inverse could not reproduce after
    /// reopening: every internal edge of `slide` and of the Parts it owns is
    /// restored as a typed owner-relative edge, so each must be spelled that
    /// way and must target a Part that exists.
    pub(crate) fn require_restorable(&self, slide: &PackURI, owned: &Owned) -> Result<()> {
        for name in std::iter::once(slide).chain(owned.parts()) {
            let node = self
                .index
                .get(&fold(name.as_str()))
                .ok_or(Error::PatchConflict)?;
            for edge in &self.nodes[*node].edges {
                if edge.target.is_none() {
                    return refusal(
                        "a relationship of the removed slide targets a missing Part, which a \
                         durable inverse cannot restore",
                    );
                }
                if !edge.canonical {
                    return refusal(
                        "primary durable removal requires canonical owner-relative \
                         relationship targets",
                    );
                }
            }
        }
        Ok(())
    }

    fn is_structural(&self, node: usize) -> bool {
        STRUCTURAL.contains(&self.nodes[node].content_type.as_str())
    }
}

struct PendingEdge {
    owner: Option<usize>,
    target: String,
    ambiguous: bool,
    canonical: bool,
}

struct Builder {
    nodes: Vec<Node>,
    index: HashMap<String, usize>,
    pending: Vec<PendingEdge>,
    max_edges: usize,
}

impl Builder {
    fn new(max_edges: usize) -> Self {
        Self {
            nodes: Vec::new(),
            index: HashMap::new(),
            pending: Vec::new(),
            max_edges,
        }
    }

    fn root(&mut self, relationships: &Relationships) -> Result<()> {
        self.edges(None, None, relationships)
    }

    fn part(
        &mut self,
        name: &PackURI,
        content_type: &str,
        relationships: &Relationships,
    ) -> Result<()> {
        let node = self.nodes.len();
        self.nodes.push(Node {
            name: name.clone(),
            content_type: content_type.to_owned(),
            edges: Vec::new(),
            incoming: Vec::new(),
        });
        self.index.insert(fold(name.as_str()), node);
        self.edges(Some(node), Some(name), relationships)
    }

    fn edges(
        &mut self,
        owner: Option<usize>,
        owner_name: Option<&PackURI>,
        relationships: &Relationships,
    ) -> Result<()> {
        for relationship in relationships.iter() {
            if relationship.is_external() {
                continue;
            }
            if self.pending.len() >= self.max_edges {
                return Err(Error::Limit {
                    kind: crate::opened::LimitKind::Codec(
                        crate::opened::CodecLimitKind::Relationships,
                    ),
                    observed: self.pending.len().checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.max_edges,
                });
            }
            let target = relationship.target_partname()?;
            let canonical = owner_name.is_none_or(|owner| {
                relationship.target_ref() == target.relative_ref(owner.base_uri())
            });
            self.pending.push(PendingEdge {
                owner,
                target: fold(target.as_str()),
                ambiguous: relationship.target_query().is_some()
                    || relationship.target_fragment().is_some(),
                canonical,
            });
        }
        Ok(())
    }

    fn finish(mut self) -> Graph {
        for edge in std::mem::take(&mut self.pending) {
            let target = self.index.get(&edge.target).copied();
            if let Some(target) = target {
                self.nodes[target].incoming.push(edge.owner);
            }
            if let Some(owner) = edge.owner {
                self.nodes[owner].edges.push(Edge {
                    target,
                    ambiguous: edge.ambiguous,
                    canonical: edge.canonical,
                });
            }
        }
        Graph {
            nodes: self.nodes,
            index: self.index,
        }
    }
}

fn fold(name: &str) -> String {
    name.to_ascii_lowercase()
}

fn refusal<T>(_detail: &str) -> Result<T> {
    Err(Error::slide_removal_refused(
        SlideRemovalRefusal::AmbiguousTopology,
    ))
}
