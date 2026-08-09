#![expect(
    clippy::assigning_clones,
    reason = "clone assignment preserves validation-before-replacement behavior"
)]
#![expect(
    clippy::manual_let_else,
    reason = "the match form keeps both parser outcomes explicit"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Source-checked glossary catalog and auxiliary-graph transactions.

use super::codec::{invalid, validate_physical_part};
use super::graph::{
    graph_matches_catalog, load_graph, package_conformance, put_graph, raw, relationship_kind,
    remove_graph, validate_raw_graph_metadata, validate_relationship_metadata,
};
use super::model::{Catalog, Conformance, Entry, Name};
use super::{OpcPackage, PackURI, Result, read, write};

/// Stable fingerprint of the complete glossary source graph.
pub type Revision = u64;

/// Immutable source snapshot of one complete glossary owner graph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    graph: Option<raw::Graph>,
    main_part: String,
    conformance: Conformance,
    revision: Revision,
}

impl Snapshot {
    /// Load and validate the glossary owner and every reachable auxiliary part.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        let main_part = package.main_document_part()?.partname().as_str().to_owned();
        let package_conformance = package_conformance(package)?;
        let graph = load_graph(package)?;
        let conformance = graph
            .as_ref()
            .map_or(package_conformance, |value| value.conformance);
        let revision = fingerprint(&main_part, conformance, graph.as_ref());
        Ok(Self {
            graph,
            main_part,
            conformance,
            revision,
        })
    }

    /// Alias for [`Self::load`] emphasizing the source-bound result.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        Self::load(package)
    }

    /// Complete raw graph, when a glossary relationship exists.
    #[must_use]
    pub fn graph(&self) -> Option<&raw::Graph> {
        self.graph.as_ref()
    }

    /// Typed catalog, when a glossary relationship exists.
    #[must_use]
    pub fn catalog(&self) -> Option<&Catalog> {
        self.graph.as_ref().map(|value| &value.catalog)
    }

    /// Entries in source order, or an empty slice for an absent owner.
    pub fn entries(&self) -> &[Entry] {
        self.catalog().map_or(&[], Catalog::entries)
    }

    /// Opaque auxiliary parts in the captured ownership closure.
    #[must_use]
    pub fn auxiliary_parts(&self) -> &[raw::Part] {
        self.graph
            .as_ref()
            .map_or(&[], |value| value.parts.as_slice())
    }

    /// Root-part relationships in the captured glossary graph.
    #[must_use]
    pub fn relationships(&self) -> &[raw::Rel] {
        self.graph
            .as_ref()
            .map_or(&[], |value| value.rels.as_slice())
    }

    /// Main document part owning the glossary relationship.
    #[must_use]
    pub fn main_part_name(&self) -> &str {
        &self.main_part
    }

    /// Conformance family used by this package or its glossary owner.
    #[must_use]
    pub const fn conformance(&self) -> Conformance {
        self.conformance
    }

    /// Exact source revision used by optimistic stale checks.
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Whether this package has no glossary owner.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.graph.is_none()
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.main_part == other.main_part
            && self.conformance == other.conformance
            && self.graph == other.graph
            && self.revision == other.revision
    }
}

/// Clone-staged, failure-atomic edit over a package-bound glossary graph.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    draft: Option<raw::Graph>,
}

impl<'a> Transaction<'a> {
    /// Capture the current package graph and begin an isolated transaction.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn new(target: &'a mut OpcPackage) -> Result<Self> {
        let before = Snapshot::load(target)?;
        Ok(Self {
            draft: before.graph.clone(),
            target,
            before,
        })
    }

    /// Immutable source snapshot used for conflict checks and inverse patches.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the currently staged typed catalog.
    #[must_use]
    pub fn catalog(&self) -> Option<&Catalog> {
        self.draft.as_ref().map(|value| &value.catalog)
    }

    /// Borrow the currently staged opaque auxiliary parts.
    #[must_use]
    pub fn auxiliary_parts(&self) -> &[raw::Part] {
        self.draft
            .as_ref()
            .map_or(&[], |value| value.parts.as_slice())
    }

    /// Borrow the currently staged root relationships.
    #[must_use]
    pub fn relationships(&self) -> &[raw::Rel] {
        self.draft
            .as_ref()
            .map_or(&[], |value| value.rels.as_slice())
    }

    /// Whether the staged complete graph differs from its source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.draft != self.before.graph
    }

    /// Replace the complete semantic catalog while retaining its graph.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn replace_catalog(&mut self, value: Catalog) -> Result<bool> {
        let changed = self.update_catalog(|catalog| {
            let changed = *catalog != value;
            *catalog = value;
            Ok(changed)
        })?;
        Ok(changed)
    }

    /// Apply an atomic mutation to the typed catalog.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn edit_catalog(
        &mut self,
        edit: impl FnOnce(&mut Catalog) -> Result<()>,
    ) -> Result<&mut Self> {
        self.update_catalog(|catalog| {
            edit(catalog)?;
            Ok(())
        })?;
        Ok(self)
    }

    /// Add one named AutoText/building-block entry.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn add_entry(&mut self, value: Entry) -> Result<usize> {
        self.update_catalog(|catalog| catalog.add(value))
    }

    /// Insert or replace one entry selected by its semantic name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn put_entry(&mut self, value: Entry) -> Result<Option<Entry>> {
        self.update_catalog(|catalog| catalog.put(value))
    }

    /// Replace one uniquely selected entry.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn replace_entry(&mut self, name: &str, value: Entry) -> Result<Option<Entry>> {
        self.update_catalog(|catalog| catalog.replace(name, value))
    }

    /// Rename one uniquely selected entry without rebuilding its body.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn rename_entry(&mut self, name: &str, value: Name) -> Result<bool> {
        self.update_catalog(|catalog| catalog.rename(name, value))
    }

    /// Remove one uniquely selected entry.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn remove_entry(&mut self, name: &str) -> Result<Option<Entry>> {
        self.update_catalog(|catalog| catalog.remove(name))
    }

    /// Reorder one entry in the semantic catalog.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn move_entry(&mut self, from: usize, to: usize) -> Result<bool> {
        self.update_catalog(|catalog| catalog.move_at(from, to))
    }

    /// Remove all entries while retaining the glossary owner and auxiliaries.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn clear_entries(&mut self) -> Result<usize> {
        self.update_catalog(|catalog| Ok(catalog.clear()))
    }

    /// Replace the root relationship set while preserving its authored IDs.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_relationships(&mut self, relationships: Vec<raw::Rel>) -> Result<bool> {
        let mut graph = self.candidate_graph();
        if graph.rels == relationships {
            return Ok(false);
        }
        graph.rels = relationships;
        validate_candidate(&graph)?;
        self.draft = Some(graph);
        Ok(true)
    }

    /// Add one root relationship, rejecting duplicate IDs atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn add_relationship(&mut self, relationship: raw::Rel) -> Result<bool> {
        let mut graph = self.candidate_graph();
        if graph.rels.iter().any(|value| value.id == relationship.id) {
            return Err(invalid("duplicate glossary relationship ID"));
        }
        validate_relationship_metadata(&relationship.id, &relationship.kind, &relationship.target)?;
        graph.rels.push(relationship);
        validate_candidate(&graph)?;
        self.draft = Some(graph);
        Ok(true)
    }

    /// Remove one root relationship by ID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn remove_relationship(&mut self, id: &str) -> Result<Option<raw::Rel>> {
        let mut graph = self.candidate_graph();
        let Some(index) = graph.rels.iter().position(|value| value.id == id) else {
            return Ok(None);
        };
        let removed = graph.rels.remove(index);
        validate_candidate(&graph)?;
        self.draft = Some(graph);
        Ok(Some(removed))
    }

    /// Add one opaque auxiliary part to the ownership graph.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn add_part(&mut self, part: raw::Part) -> Result<bool> {
        let mut graph = self.candidate_graph();
        let uri = PackURI::new(&part.name).map_err(crate::Error::Uri)?;
        if uri.is_equivalent_to(&PackURI::new(&graph.root_name).map_err(crate::Error::Uri)?)
            || graph.parts.iter().any(|value| {
                PackURI::new(&value.name)
                    .ok()
                    .is_some_and(|candidate| candidate.is_equivalent_to(&uri))
            })
        {
            return Err(invalid(format!(
                "glossary auxiliary part '{uri}' already exists"
            )));
        }
        graph.parts.push(part);
        validate_candidate(&graph)?;
        self.draft = Some(graph);
        Ok(true)
    }

    /// Replace one opaque auxiliary part by its absolute part name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn replace_part(
        &mut self,
        name: &str,
        content_type: impl Into<String>,
        data: Vec<u8>,
        relationships: Vec<raw::Rel>,
    ) -> Result<bool> {
        let mut part = raw::Part::new(name, content_type, data)?;
        part.set_relationships(relationships)?;
        let mut graph = self.candidate_graph();
        let requested = PackURI::new(name).map_err(crate::Error::Uri)?;
        let index = graph.parts.iter().position(|value| {
            PackURI::new(&value.name)
                .ok()
                .is_some_and(|candidate| candidate.is_equivalent_to(&requested))
        });
        let Some(index) = index else {
            return Err(invalid(format!(
                "glossary auxiliary part '{requested}' is absent"
            )));
        };
        if graph.parts[index] == part {
            return Ok(false);
        }
        part.name = graph.parts[index].name.clone();
        graph.parts[index] = part;
        validate_candidate(&graph)?;
        self.draft = Some(graph);
        Ok(true)
    }

    /// Replace one opaque auxiliary payload while retaining its relationships.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn replace_part_data(&mut self, name: &str, data: Vec<u8>) -> Result<bool> {
        let mut graph = self.candidate_graph();
        let index = part_index(&graph, name)?;
        let part = &mut graph.parts[index];
        if part.data() == data.as_slice() {
            return Ok(false);
        }
        part.replace_data(data)?;
        validate_candidate(&graph)?;
        self.draft = Some(graph);
        Ok(true)
    }

    /// Replace one opaque auxiliary relationship set.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_part_relationships(
        &mut self,
        name: &str,
        relationships: Vec<raw::Rel>,
    ) -> Result<bool> {
        let mut graph = self.candidate_graph();
        let index = part_index(&graph, name)?;
        if graph.parts[index].relationships() == relationships.as_slice() {
            return Ok(false);
        }
        graph.parts[index].set_relationships(relationships)?;
        validate_candidate(&graph)?;
        self.draft = Some(graph);
        Ok(true)
    }

    /// Remove one auxiliary part; dangling relationships are rejected at commit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn remove_part(&mut self, name: &str) -> Result<Option<raw::Part>> {
        let mut graph = self.candidate_graph();
        let index = match part_index(&graph, name) {
            Ok(index) => index,
            Err(_) => return Ok(None),
        };
        let removed = graph.parts.remove(index);
        validate_candidate(&graph)?;
        self.draft = Some(graph);
        Ok(Some(removed))
    }

    /// Remove the complete glossary owner and its exclusively owned graph.
    pub fn remove_glossary(&mut self) -> bool {
        self.draft.take().is_some()
    }

    /// Validate and publish the staged graph atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        let current = Snapshot::load(self.target)?;
        if !current.same_source(&self.before) {
            return Err(invalid("glossary transaction source is stale"));
        }
        let mut candidate = self.target.clone();
        match self.draft.as_ref() {
            Some(graph) => {
                put_graph(&mut candidate, graph)?;
            },
            None => {
                remove_graph(&mut candidate)?;
            },
        }
        let snapshot = Snapshot::load(&candidate)?;
        match (self.draft.as_ref(), snapshot.graph.as_ref()) {
            (Some(expected), Some(actual)) => {
                let expected_catalog = published_catalog(expected)?;
                if !graph_matches_catalog(actual, expected, &expected_catalog, None)? {
                    return Err(invalid("glossary publication changed the staged graph"));
                }
            },
            (None, None) => {},
            _ => return Err(invalid("glossary publication changed the staged graph")),
        }
        let patch = Patch::new(self.before, snapshot.clone());
        *self.target = candidate;
        Ok(Commit::new(snapshot, patch, true))
    }

    fn candidate_graph(&self) -> raw::Graph {
        self.draft
            .clone()
            .unwrap_or_else(|| raw::Graph::new(Catalog::new(), self.before.conformance))
    }

    fn update_catalog<T>(&mut self, edit: impl FnOnce(&mut Catalog) -> Result<T>) -> Result<T> {
        let mut graph = self.candidate_graph();
        let result = edit(&mut graph.catalog)?;
        graph.catalog.validate_bound_lineages()?;
        write(&graph.catalog, self.before.conformance)?;
        self.draft = Some(graph);
        Ok(result)
    }
}

/// Successful glossary publication and its reversible source patch.
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

    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Reversible, source-checked replacement of the complete glossary graph.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply the patch atomically after checking the complete owner graph.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<Snapshot> {
        let current = Snapshot::load(target)?;
        if !current.same_source(&self.before) {
            return Err(invalid("glossary patch source is stale"));
        }
        if self.is_empty() {
            return Ok(current);
        }
        let mut candidate = target.clone();
        match self.after.graph.as_ref() {
            Some(graph) => {
                put_graph(&mut candidate, graph)?;
            },
            None => {
                remove_graph(&mut candidate)?;
            },
        }
        let resulting = Snapshot::load(&candidate)?;
        if !resulting.same_source(&self.after) {
            return Err(invalid(
                "glossary patch publication changed its target graph",
            ));
        }
        *target = candidate;
        Ok(resulting)
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

fn part_index(graph: &raw::Graph, name: &str) -> Result<usize> {
    let requested = PackURI::new(name).map_err(crate::Error::Uri)?;
    graph
        .parts
        .iter()
        .position(|value| {
            PackURI::new(&value.name)
                .ok()
                .is_some_and(|candidate| candidate.is_equivalent_to(&requested))
        })
        .ok_or_else(|| invalid(format!("glossary auxiliary part '{requested}' is absent")))
}

fn validate_candidate(graph: &raw::Graph) -> Result<()> {
    validate_raw_graph_metadata(graph)?;
    write(&graph.catalog, graph.conformance)?;
    for part in &graph.parts {
        validate_physical_part(&part.name, &part.content_type, part.data().len())?;
    }
    for relationship in graph.rels.iter().chain(
        graph
            .parts
            .iter()
            .flat_map(|part| part.relationships().iter()),
    ) {
        relationship_kind(graph.conformance, &relationship.kind).ok_or_else(|| {
            invalid(format!(
                "unsupported glossary relationship type '{}' for {:?} conformance",
                relationship.kind, graph.conformance
            ))
        })?;
    }
    Ok(())
}

fn published_catalog(graph: &raw::Graph) -> Result<Catalog> {
    if let Some(xml) = &graph.root_xml {
        let (catalog, conformance) = read(xml.as_slice())?;
        if conformance == graph.conformance && catalog == graph.catalog {
            return Ok(graph.catalog.clone());
        }
    }
    let encoded = write(&graph.catalog, graph.conformance)?;
    let (catalog, conformance) = read(&encoded)?;
    if conformance != graph.conformance {
        return Err(invalid("canonical glossary changed conformance"));
    }
    Ok(catalog)
}

fn fingerprint(main_part: &str, conformance: Conformance, graph: Option<&raw::Graph>) -> Revision {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    feed_text(&mut hash, main_part);
    feed_text(
        &mut hash,
        if conformance == Conformance::Strict {
            "strict"
        } else {
            "transitional"
        },
    );
    let Some(graph) = graph else {
        hash ^= 0;
        return hash;
    };
    feed_text(&mut hash, &graph.root_name);
    if let Some(value) = &graph.root_xml {
        feed_bytes(&mut hash, value);
    }
    for relationship in &graph.rels {
        feed_text(&mut hash, &relationship.id);
        feed_text(&mut hash, &relationship.kind);
        feed_text(&mut hash, &relationship.target);
        hash ^= u64::from(relationship.external);
    }
    for part in &graph.parts {
        feed_text(&mut hash, &part.name);
        feed_text(&mut hash, &part.content_type);
        feed_bytes(&mut hash, part.data());
        for relationship in &part.rels {
            feed_text(&mut hash, &relationship.id);
            feed_text(&mut hash, &relationship.kind);
            feed_text(&mut hash, &relationship.target);
            hash ^= u64::from(relationship.external);
        }
    }
    hash
}

fn feed_text(hash: &mut u64, value: &str) {
    feed_bytes(hash, value.as_bytes());
}

fn feed_bytes(hash: &mut u64, value: &[u8]) {
    for byte in value {
        *hash ^= u64::from(*byte);
        *hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    *hash ^= value.len() as u64;
    *hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
}
