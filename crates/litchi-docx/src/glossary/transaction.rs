#![expect(
    clippy::manual_let_else,
    reason = "the match form keeps both parser outcomes explicit"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Source-checked glossary catalog and auxiliary-graph transactions.

use super::codec::read_with_limits;
use super::codec::{
    WritePlan, invalid, name_key_with_limits, plan_write_with_limits, validate_authored_name,
    validate_entry_limits, validate_physical_part, write_with_limits,
};
use super::graph::{
    graph_matches_catalog, package_conformance, prepare_graph_for_put, put_prepared_graph, raw,
    relationship_kind, remove_graph_with_limits, validate_raw_graph_metadata,
    validate_relationship_metadata,
};
use super::model::{Catalog, Conformance, Entry, Name};
use super::{CT, MAX_VALUES, OpcPackage, PackURI, Result};

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
    /// The required raw policy is checked before the graph is allocated.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn load(package: &OpcPackage, limits: &super::Limits) -> Result<Self> {
        let package_conformance = package_conformance(package)?;
        let graph = raw::load(package, limits)?;
        let main_part = package.main_document_part()?.partname().as_str().to_owned();
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
    /// The required raw policy is checked before the graph is allocated.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn read(package: &OpcPackage, limits: &super::Limits) -> Result<Self> {
        Self::load(package, limits)
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
#[derive(Clone, Debug)]
pub(crate) struct Transaction {
    before: Snapshot,
    draft: Option<raw::Graph>,
    limits: super::Limits,
    empty_seed: EmptyGraphSeed,
    empty_usage: Option<GraphUsage>,
}

impl Transaction {
    /// Capture the current package graph and begin an isolated transaction.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn new(target: &OpcPackage) -> Result<Self> {
        Self::new_with_limits(target, &super::Limits::DEFAULT)
    }

    pub(in crate::glossary) fn new_with_limits(
        target: &OpcPackage,
        limits: &super::Limits,
    ) -> Result<Self> {
        let before = Snapshot::load(target, limits)?;
        let empty_seed = empty_graph_seed(target, before.conformance)?;
        Ok(Self {
            draft: before.graph.clone(),
            before,
            limits: *limits,
            empty_seed,
            empty_usage: None,
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
        let limits = self.limits;
        let changed = self.update_catalog(|catalog| {
            let changed = *catalog != value;
            *catalog = value;
            catalog.bind_edit_limits(&limits);
            catalog.rebuild_state(&limits)?;
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
        let limits = self.limits;
        validate_authored_name(&value)?;
        validate_entry_limits(&value, &limits)?;
        self.update_catalog(|catalog| catalog.add_with_limits(value, &limits))
    }

    /// Insert or replace one entry selected by its semantic name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn put_entry(&mut self, value: Entry) -> Result<Option<Entry>> {
        let limits = self.limits;
        validate_authored_name(&value)?;
        validate_entry_limits(&value, &limits)?;
        self.update_catalog(|catalog| catalog.put_with_limits(value, &limits))
    }

    /// Replace one uniquely selected entry.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn replace_entry(&mut self, name: &str, value: Entry) -> Result<Option<Entry>> {
        let limits = self.limits;
        name_key_with_limits(name, &limits)?;
        validate_authored_name(&value)?;
        validate_entry_limits(&value, &limits)?;
        self.update_catalog(|catalog| catalog.replace_with_limits(name, value, &limits))
    }

    /// Rename one uniquely selected entry without rebuilding its body.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn rename_entry(&mut self, name: &str, value: Name) -> Result<bool> {
        let limits = self.limits;
        self.update_catalog(|catalog| catalog.rename_with_limits(name, value, &limits))
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
        let unchanged = self
            .draft
            .as_ref()
            .is_some_and(|graph| graph.rels == relationships);
        if unchanged {
            return Ok(false);
        }
        let empty_usage = self.empty_usage_for_current()?;
        preflight_root_relationships(
            self.draft.as_ref(),
            &empty_usage,
            &relationships,
            &self.limits,
        )?;
        let mut graph = self.candidate_graph();
        graph.rels = relationships;
        validate_candidate(&graph, &self.limits)?;
        self.draft = Some(graph);
        Ok(true)
    }

    /// Add one root relationship, rejecting duplicate IDs atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn add_relationship(&mut self, relationship: raw::Rel) -> Result<bool> {
        let empty_usage = self.empty_usage_for_current()?;
        preflight_add_relationship(
            self.draft.as_ref(),
            &empty_usage,
            &relationship,
            &self.limits,
        )?;
        let mut graph = self.candidate_graph();
        graph.rels.push(relationship);
        validate_candidate(&graph, &self.limits)?;
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
        validate_candidate(&graph, &self.limits)?;
        self.draft = Some(graph);
        Ok(Some(removed))
    }

    /// Add one opaque auxiliary part to the ownership graph.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn add_part(&mut self, part: raw::Part) -> Result<bool> {
        let empty_usage = self.empty_usage_for_current()?;
        preflight_add_part(self.draft.as_ref(), &empty_usage, &part, &self.limits)?;
        let mut graph = self.candidate_graph();
        graph.parts.push(part);
        validate_candidate(&graph, &self.limits)?;
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
        content_type: impl AsRef<str>,
        data: Vec<u8>,
        relationships: Vec<raw::Rel>,
    ) -> Result<bool> {
        let content_type = content_type.as_ref();
        preflight_part_name_and_type(name, content_type, &self.limits)?;
        let requested = PackURI::new(name).map_err(crate::Error::Uri)?;
        let graph = self
            .draft
            .as_ref()
            .ok_or_else(|| invalid(format!("glossary auxiliary part '{requested}' is absent")))?;
        let old_index = part_index(graph, name)?;
        let old = graph
            .parts
            .get(old_index)
            .ok_or_else(|| invalid("glossary auxiliary part index changed"))?;
        if old.content_type == content_type
            && old.data() == data.as_slice()
            && old.relationships() == relationships.as_slice()
        {
            return Ok(false);
        }
        preflight_replace_part_input(
            graph,
            old_index,
            name,
            content_type,
            &data,
            &relationships,
            &self.limits,
        )?;
        let mut part = raw::Part::new(name, content_type, data, &self.limits)?;
        part.set_relationships(relationships, &self.limits)?;
        let old_name = old.name.clone();
        part.name = old_name;
        let mut graph = self.candidate_graph();
        graph.parts[old_index] = part;
        validate_candidate(&graph, &self.limits)?;
        self.draft = Some(graph);
        Ok(true)
    }

    /// Replace one opaque auxiliary payload while retaining its relationships.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn replace_part_data(&mut self, name: &str, data: Vec<u8>) -> Result<bool> {
        let graph = self
            .draft
            .as_ref()
            .ok_or_else(|| invalid(format!("glossary auxiliary part '{name}' is absent")))?;
        let index = part_index(graph, name)?;
        let part = graph
            .parts
            .get(index)
            .ok_or_else(|| invalid("glossary auxiliary part index changed"))?;
        if part.data() == data.as_slice() {
            return Ok(false);
        }
        preflight_replace_part_data(graph, index, &data, &self.limits)?;
        let mut graph = self.candidate_graph();
        graph.parts[index].replace_data(data, &self.limits)?;
        validate_candidate(&graph, &self.limits)?;
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
        let current = self
            .draft
            .as_ref()
            .ok_or_else(|| invalid(format!("glossary auxiliary part '{name}' is absent")))?;
        let index = part_index(current, name)?;
        if current.parts[index].relationships() == relationships.as_slice() {
            return Ok(false);
        }
        preflight_part_relationships(current, index, &relationships, &self.limits)?;
        let mut graph = self.candidate_graph();
        graph.parts[index].set_relationships(relationships, &self.limits)?;
        validate_candidate(&graph, &self.limits)?;
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
        validate_candidate(&graph, &self.limits)?;
        self.draft = Some(graph);
        Ok(Some(removed))
    }

    /// Remove the complete glossary owner and its exclusively owned graph.
    pub fn remove_glossary(&mut self) -> bool {
        self.draft.take().is_some()
    }

    /// Validate and publish the staged graph atomically.
    ///
    /// A transaction without a change publishes nothing, so a signed package
    /// keeps its signature byte for byte. A changed commit over a signed
    /// package is refused with
    /// [`litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy`] and leaves
    /// it untouched (ADR 0006): remove the signature explicitly first
    /// ([`OpcPackage::unsign`]).
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn commit(self, target: &mut OpcPackage) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        crate::signature_policy::ensure_unsigned(target)?;
        let prepared = self
            .draft
            .as_ref()
            .map(|graph| prepare_graph_for_put(target, graph, &self.limits))
            .transpose()?;
        let current = Snapshot::load(target, &self.limits)?;
        if !current.same_source(&self.before) {
            return Err(invalid("glossary transaction source is stale"));
        }
        let mut candidate = target.clone();
        match (self.draft.as_ref(), prepared) {
            (Some(graph), Some(prepared)) => {
                put_prepared_graph(&mut candidate, graph, &self.limits, prepared)?;
            },
            (None, None) => {
                remove_graph_with_limits(&mut candidate, &self.limits)?;
            },
            _ => return Err(invalid("glossary transaction draft preparation changed")),
        }
        let snapshot = Snapshot::load(&candidate, &self.limits)?;
        match (self.draft.as_ref(), snapshot.graph.as_ref()) {
            (Some(expected), Some(actual)) => {
                let expected_catalog = published_catalog(expected, &self.limits)?;
                if !graph_matches_catalog(actual, expected, &expected_catalog, None)? {
                    return Err(invalid("glossary publication changed the staged graph"));
                }
            },
            (None, None) => {},
            _ => return Err(invalid("glossary publication changed the staged graph")),
        }
        let patch = Patch::new(self.before, snapshot.clone());
        *target = candidate;
        Ok(Commit::new(snapshot, patch, true))
    }

    fn candidate_graph(&self) -> raw::Graph {
        let mut graph = self
            .draft
            .clone()
            .unwrap_or_else(|| raw::Graph::new(Catalog::new(), self.before.conformance));
        if graph.owner_main.is_none() && graph.owner_id.is_none() && graph.owner_target.is_none() {
            graph.owner_main = Some(self.empty_seed.main.clone());
            graph.owner_id = Some(self.empty_seed.owner_id.clone());
            graph.owner_target = Some(self.empty_seed.owner_target.clone());
        }
        graph.catalog.bind_edit_limits(&self.limits);
        graph
    }

    fn empty_usage_for_current(&mut self) -> Result<GraphUsage> {
        if self.draft.is_some() {
            return Ok(GraphUsage::default());
        }
        if let Some(usage) = self.empty_usage {
            return Ok(usage);
        }
        let usage = generated_graph_usage(&self.empty_seed, &self.limits)?;
        self.empty_usage = Some(usage);
        Ok(usage)
    }

    fn update_catalog<T>(&mut self, edit: impl FnOnce(&mut Catalog) -> Result<T>) -> Result<T> {
        let mut graph = self.candidate_graph();
        let result = edit(&mut graph.catalog)?;
        graph.catalog.validate_bound_lineages()?;
        graph.catalog.validate_for_limits(&self.limits)?;
        let plan = plan_write_with_limits(&graph.catalog, graph.conformance, &self.limits)?;
        if plan.bytes > self.limits.xml_bytes() {
            return Err(crate::Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::XmlBytes,
                actual: plan.bytes,
                maximum: self.limits.xml_bytes(),
            });
        }
        if self
            .draft
            .as_ref()
            .is_none_or(|previous| previous.catalog != graph.catalog)
        {
            graph.root_xml = None;
        }
        check_usage(
            graph_usage_with_plan(&graph, &self.limits, Some(&plan))?,
            &self.limits,
        )?;
        self.draft = Some(graph);
        Ok(result)
    }
}

/// Successful glossary publication and its reversible source patch.
#[derive(Clone, Debug)]
pub(crate) struct Commit {
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
pub(crate) struct Patch {
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
    /// An empty patch publishes nothing and keeps a signed target exactly as
    /// it was. Any other patch changes the target, so a signed target is
    /// refused with [`litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy`]
    /// (ADR 0006): remove its signature explicitly first
    /// ([`OpcPackage::unsign`]). The inverse never restores a signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn apply(&self, target: &mut OpcPackage, limits: &super::Limits) -> Result<Snapshot> {
        let prepared = if self.is_empty() {
            None
        } else {
            self.after
                .graph
                .as_ref()
                .map(|graph| prepare_graph_for_put(target, graph, limits))
                .transpose()?
        };
        let current = Snapshot::load(target, limits)?;
        if !current.same_source(&self.before) {
            return Err(invalid("glossary patch source is stale"));
        }
        if self.is_empty() {
            return Ok(current);
        }
        crate::signature_policy::ensure_unsigned(target)?;
        let mut candidate = target.clone();
        match (self.after.graph.as_ref(), prepared) {
            (Some(graph), Some(prepared)) => {
                put_prepared_graph(&mut candidate, graph, limits, prepared)?;
            },
            (None, None) => {
                remove_graph_with_limits(&mut candidate, limits)?;
            },
            _ => return Err(invalid("glossary patch preparation changed")),
        }
        let resulting = Snapshot::load(&candidate, limits)?;
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
    pub fn undo(&self, target: &mut OpcPackage, limits: &super::Limits) -> Result<Snapshot> {
        self.inverse().apply(target, limits)
    }
}

const DEFAULT_ROOT_NAME: &str = "/word/glossary/document.xml";

#[derive(Clone, Copy, Debug, Default)]
struct GraphUsage {
    bytes: usize,
    metadata: usize,
    relationships: usize,
    parts: usize,
}

#[derive(Clone, Debug)]
struct EmptyGraphSeed {
    root: String,
    main: String,
    owner_id: String,
    owner_target: String,
    conformance: Conformance,
}

fn graph_usage(graph: &raw::Graph, limits: &super::Limits) -> Result<GraphUsage> {
    graph_usage_with_plan(graph, limits, None)
}

fn graph_usage_with_plan(
    graph: &raw::Graph,
    limits: &super::Limits,
    root_plan: Option<&WritePlan>,
) -> Result<GraphUsage> {
    let mut usage = GraphUsage {
        metadata: graph.root_name.len(),
        relationships: graph.rels.len(),
        parts: graph.parts.len(),
        ..GraphUsage::default()
    };
    add_usage_metadata(&mut usage.metadata, &graph.root_name)?;
    add_usage_metadata(&mut usage.metadata, CT)?;
    let root_bytes = if let Some(xml) = graph.root_xml() {
        xml.len()
    } else if let Some(plan) = root_plan {
        plan.bytes
    } else {
        plan_write_with_limits(&graph.catalog, graph.conformance, limits)?.bytes
    };
    add_usage_bytes(&mut usage.bytes, root_bytes)?;
    for value in [
        graph.owner_main.as_deref(),
        graph.owner_id.as_deref(),
        graph.owner_target.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        add_usage_metadata(&mut usage.metadata, value)?;
    }
    for relationship in &graph.rels {
        add_relationship_metadata(&mut usage, relationship)?;
    }
    for part in &graph.parts {
        add_usage_bytes(&mut usage.bytes, part.data().len())?;
        add_usage_metadata(&mut usage.metadata, &part.name)?;
        add_usage_metadata(&mut usage.metadata, &part.content_type)?;
        usage.relationships = usage
            .relationships
            .checked_add(part.relationships().len())
            .ok_or_else(|| invalid("glossary relationship count overflow"))?;
        for relationship in part.relationships() {
            add_relationship_metadata(&mut usage, relationship)?;
        }
    }
    Ok(usage)
}

fn empty_graph_seed(package: &OpcPackage, conformance: Conformance) -> Result<EmptyGraphSeed> {
    let root = PackURI::new(DEFAULT_ROOT_NAME).map_err(crate::Error::Uri)?;
    let main = package.main_document_part()?;
    Ok(EmptyGraphSeed {
        root: root.as_str().to_owned(),
        main: main.partname().as_str().to_owned(),
        owner_id: next_relationship_id(main)?,
        owner_target: root.relative_ref(main.partname().base_uri()),
        conformance,
    })
}

fn generated_graph_usage(seed: &EmptyGraphSeed, limits: &super::Limits) -> Result<GraphUsage> {
    let root_xml = write_with_limits(&Catalog::new(), seed.conformance, limits)?;
    let mut usage = GraphUsage {
        bytes: root_xml.len(),
        ..GraphUsage::default()
    };
    selected_metadata(&mut usage, &seed.root, true, limits)?;
    selected_metadata(&mut usage, &seed.root, true, limits)?;
    selected_metadata(&mut usage, CT, false, limits)?;
    selected_metadata(&mut usage, &seed.main, true, limits)?;
    selected_metadata(&mut usage, &seed.owner_id, false, limits)?;
    selected_metadata(&mut usage, &seed.owner_target, false, limits)?;
    check_usage(usage, limits)?;
    Ok(usage)
}

fn next_relationship_id(part: &dyn super::Part) -> Result<String> {
    next_relationship_id_from(part, 1)
}

fn next_relationship_id_from(part: &dyn super::Part, start: u32) -> Result<String> {
    let mut next = start;
    loop {
        let used = part.rels().iter().any(|relationship| {
            relationship
                .r_id()
                .strip_prefix("rId")
                .and_then(|value| atoi_simd::parse::<u32, false, false>(value.as_bytes()).ok())
                == Some(next)
        });
        if !used {
            return Ok(format!("rId{next}"));
        }
        next = next
            .checked_add(1)
            .ok_or_else(|| invalid("no glossary relationship ID is available"))?;
    }
}

fn staged_graph_usage(
    draft: Option<&raw::Graph>,
    empty_usage: &GraphUsage,
    limits: &super::Limits,
) -> Result<GraphUsage> {
    let usage = draft.map_or_else(|| Ok(*empty_usage), |graph| graph_usage(graph, limits))?;
    let root_name = draft.map_or(DEFAULT_ROOT_NAME, |graph| graph.root_name.as_str());
    if root_name.len() > limits.name_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::StringBytes,
            actual: root_name.len(),
            maximum: limits.name_bytes(),
        });
    }
    check_usage(usage, limits)?;
    Ok(usage)
}

fn add_usage_bytes(total: &mut usize, amount: usize) -> Result<()> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| invalid("glossary graph byte count overflow"))?;
    Ok(())
}

fn add_usage_metadata(total: &mut usize, value: &str) -> Result<()> {
    *total = total
        .checked_add(value.len())
        .ok_or_else(|| invalid("glossary graph metadata size overflow"))?;
    Ok(())
}

fn add_relationship_metadata(usage: &mut GraphUsage, relationship: &raw::Rel) -> Result<()> {
    for value in [&relationship.id, &relationship.kind, &relationship.target] {
        add_usage_metadata(&mut usage.metadata, value)?;
    }
    Ok(())
}

fn check_usage(usage: GraphUsage, limits: &super::Limits) -> Result<()> {
    if usage.bytes > limits.graph_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::GraphBytes,
            actual: usage.bytes,
            maximum: limits.graph_bytes(),
        });
    }
    if usage.metadata > limits.metadata_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::GraphMetadataBytes,
            actual: usage.metadata,
            maximum: limits.metadata_bytes(),
        });
    }
    let relationship_limit = limits.values().min(MAX_VALUES);
    if usage.relationships > relationship_limit {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::GraphWideRelationship,
            actual: usage.relationships,
            maximum: relationship_limit,
        });
    }
    let part_limit = limits.parts().min(MAX_VALUES);
    if usage.parts > part_limit {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::AuxiliaryPart,
            actual: usage.parts,
            maximum: part_limit,
        });
    }
    Ok(())
}

fn selected_metadata(
    usage: &mut GraphUsage,
    value: &str,
    name: bool,
    limits: &super::Limits,
) -> Result<()> {
    let maximum = if name {
        limits.name_bytes()
    } else {
        limits.string_bytes()
    };
    if value.len() > maximum {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::StringBytes,
            actual: value.len(),
            maximum,
        });
    }
    add_usage_metadata(&mut usage.metadata, value)
}

fn relationship_usage(relationships: &[raw::Rel], limits: &super::Limits) -> Result<GraphUsage> {
    if relationships.len() > limits.values() {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::Relationship,
            actual: relationships.len(),
            maximum: limits.values(),
        });
    }
    let mut usage = GraphUsage {
        relationships: relationships.len(),
        ..GraphUsage::default()
    };
    for relationship in relationships {
        validate_relationship_metadata(&relationship.id, &relationship.kind, &relationship.target)?;
        for value in [&relationship.id, &relationship.kind, &relationship.target] {
            selected_metadata(&mut usage, value, false, limits)?;
        }
    }
    Ok(usage)
}

fn replace_relationship_usage(
    mut usage: GraphUsage,
    old: &[raw::Rel],
    new: &[raw::Rel],
    limits: &super::Limits,
) -> Result<()> {
    let old_usage = relationship_usage(old, limits)?;
    let new_usage = relationship_usage(new, limits)?;
    usage.relationships = usage
        .relationships
        .checked_sub(old_usage.relationships)
        .and_then(|value| value.checked_add(new_usage.relationships))
        .ok_or_else(|| invalid("glossary relationship count underflow"))?;
    usage.metadata = usage
        .metadata
        .checked_sub(old_usage.metadata)
        .and_then(|value| value.checked_add(new_usage.metadata))
        .ok_or_else(|| invalid("glossary graph metadata size underflow"))?;
    check_usage(usage, limits)
}

fn preflight_root_relationships(
    draft: Option<&raw::Graph>,
    empty_usage: &GraphUsage,
    relationships: &[raw::Rel],
    limits: &super::Limits,
) -> Result<()> {
    let usage = staged_graph_usage(draft, empty_usage, limits)?;
    let old = draft.map_or(&[][..], |graph| graph.rels.as_slice());
    replace_relationship_usage(usage, old, relationships, limits)
}

fn preflight_add_relationship(
    draft: Option<&raw::Graph>,
    empty_usage: &GraphUsage,
    relationship: &raw::Rel,
    limits: &super::Limits,
) -> Result<()> {
    let Some(graph) = draft else {
        let usage = staged_graph_usage(None, empty_usage, limits)?;
        let new = relationship_usage(std::slice::from_ref(relationship), limits)?;
        let usage = GraphUsage {
            relationships: usage
                .relationships
                .checked_add(new.relationships)
                .ok_or_else(|| invalid("glossary relationship count overflow"))?,
            metadata: usage
                .metadata
                .checked_add(new.metadata)
                .ok_or_else(|| invalid("glossary graph metadata size overflow"))?,
            ..usage
        };
        return check_usage(usage, limits);
    };
    if graph.rels.iter().any(|value| value.id == relationship.id) {
        return Err(invalid("duplicate glossary relationship ID"));
    }
    let usage = staged_graph_usage(Some(graph), empty_usage, limits)?;
    let new = relationship_usage(std::slice::from_ref(relationship), limits)?;
    let usage = GraphUsage {
        relationships: usage
            .relationships
            .checked_add(new.relationships)
            .ok_or_else(|| invalid("glossary relationship count overflow"))?,
        metadata: usage
            .metadata
            .checked_add(new.metadata)
            .ok_or_else(|| invalid("glossary graph metadata size overflow"))?,
        ..usage
    };
    check_usage(usage, limits)
}

fn part_usage_with_data(
    part: &raw::Part,
    data: &[u8],
    limits: &super::Limits,
    admit_payload: bool,
) -> Result<GraphUsage> {
    part_usage_values(
        &part.name,
        &part.content_type,
        data,
        part.relationships(),
        limits,
        admit_payload,
    )
}

fn part_usage_values(
    name: &str,
    content_type: &str,
    data: &[u8],
    relationships: &[raw::Rel],
    limits: &super::Limits,
    admit_payload: bool,
) -> Result<GraphUsage> {
    let mut usage = GraphUsage {
        bytes: data.len(),
        parts: 1,
        ..GraphUsage::default()
    };
    selected_metadata(&mut usage, name, true, limits)?;
    selected_metadata(&mut usage, content_type, false, limits)?;
    check_usage(usage, limits)?;
    validate_physical_part(name, content_type, data.len())?;
    let relationship_usage = relationship_usage(relationships, limits)?;
    usage.metadata = usage
        .metadata
        .checked_add(relationship_usage.metadata)
        .ok_or_else(|| invalid("glossary graph metadata size overflow"))?;
    usage.relationships = relationship_usage.relationships;
    if admit_payload && is_xml_content_type(content_type) {
        raw::admit_xml(data, limits)?;
    }
    check_usage(usage, limits)?;
    Ok(usage)
}

fn is_xml_content_type(content_type: &str) -> bool {
    content_type.ends_with("+xml")
        || content_type.eq_ignore_ascii_case("application/xml")
        || content_type.eq_ignore_ascii_case("text/xml")
}

fn preflight_add_part(
    draft: Option<&raw::Graph>,
    empty_usage: &GraphUsage,
    part: &raw::Part,
    limits: &super::Limits,
) -> Result<()> {
    let usage = staged_graph_usage(draft, empty_usage, limits)?;
    let root_name = draft.map_or(DEFAULT_ROOT_NAME, |graph| graph.root_name.as_str());
    let uri = PackURI::new(&part.name).map_err(crate::Error::Uri)?;
    if uri.is_equivalent_to(&PackURI::new(root_name).map_err(crate::Error::Uri)?)
        || draft.is_some_and(|graph| {
            graph.parts.iter().any(|value| {
                PackURI::new(&value.name)
                    .ok()
                    .is_some_and(|candidate| candidate.is_equivalent_to(&uri))
            })
        })
    {
        return Err(invalid(format!(
            "glossary auxiliary part '{uri}' already exists"
        )));
    }
    let new = part_usage_with_data(part, part.data(), limits, true)?;
    let usage = GraphUsage {
        bytes: usage
            .bytes
            .checked_add(new.bytes)
            .ok_or_else(|| invalid("glossary graph byte count overflow"))?,
        metadata: usage
            .metadata
            .checked_add(new.metadata)
            .ok_or_else(|| invalid("glossary graph metadata size overflow"))?,
        relationships: usage
            .relationships
            .checked_add(new.relationships)
            .ok_or_else(|| invalid("glossary relationship count overflow"))?,
        parts: usage
            .parts
            .checked_add(1)
            .ok_or_else(|| invalid("glossary auxiliary-part count overflow"))?,
    };
    check_usage(usage, limits)
}

fn preflight_replace_part_input(
    graph: &raw::Graph,
    index: usize,
    requested_name: &str,
    content_type: &str,
    data: &[u8],
    relationships: &[raw::Rel],
    limits: &super::Limits,
) -> Result<()> {
    let old = graph
        .parts
        .get(index)
        .ok_or_else(|| invalid("glossary auxiliary part index changed"))?;
    // Reproduce `raw::Part::new` and `set_relationships` admission while the
    // caller still owns the input buffers.  The requested name is checked as
    // authored, but the aggregate ledger uses the existing canonical name
    // retained by replacement.
    part_usage_values(
        requested_name,
        content_type,
        data,
        relationships,
        limits,
        true,
    )?;
    let mut usage = staged_graph_usage(Some(graph), &GraphUsage::default(), limits)?;
    let old_usage = part_usage_with_data(old, old.data(), limits, false)?;
    let new_usage = part_usage_values(&old.name, content_type, data, relationships, limits, false)?;
    usage.bytes = usage
        .bytes
        .checked_sub(old_usage.bytes)
        .and_then(|value| value.checked_add(new_usage.bytes))
        .ok_or_else(|| invalid("glossary graph byte count underflow"))?;
    usage.metadata = usage
        .metadata
        .checked_sub(old_usage.metadata)
        .and_then(|value| value.checked_add(new_usage.metadata))
        .ok_or_else(|| invalid("glossary graph metadata size underflow"))?;
    usage.relationships = usage
        .relationships
        .checked_sub(old_usage.relationships)
        .and_then(|value| value.checked_add(new_usage.relationships))
        .ok_or_else(|| invalid("glossary relationship count underflow"))?;
    check_usage(usage, limits)
}

fn preflight_part_name_and_type(
    name: &str,
    content_type: &str,
    limits: &super::Limits,
) -> Result<()> {
    if name.len() > limits.name_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::StringBytes,
            actual: name.len(),
            maximum: limits.name_bytes(),
        });
    }
    if content_type.len() > limits.string_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::StringBytes,
            actual: content_type.len(),
            maximum: limits.string_bytes(),
        });
    }
    Ok(())
}

fn preflight_replace_part_data(
    graph: &raw::Graph,
    index: usize,
    data: &[u8],
    limits: &super::Limits,
) -> Result<()> {
    let old = graph
        .parts
        .get(index)
        .ok_or_else(|| invalid("glossary auxiliary part index changed"))?;
    let mut usage = staged_graph_usage(Some(graph), &GraphUsage::default(), limits)?;
    let old_usage = part_usage_with_data(old, old.data(), limits, false)?;
    let new_usage = part_usage_with_data(old, data, limits, true)?;
    usage.bytes = usage
        .bytes
        .checked_sub(old_usage.bytes)
        .and_then(|value| value.checked_add(new_usage.bytes))
        .ok_or_else(|| invalid("glossary graph byte count underflow"))?;
    check_usage(usage, limits)
}

fn preflight_part_relationships(
    graph: &raw::Graph,
    index: usize,
    relationships: &[raw::Rel],
    limits: &super::Limits,
) -> Result<()> {
    let part = graph
        .parts
        .get(index)
        .ok_or_else(|| invalid("glossary auxiliary part index changed"))?;
    let usage = staged_graph_usage(Some(graph), &GraphUsage::default(), limits)?;
    replace_relationship_usage(usage, part.relationships(), relationships, limits)
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

fn validate_candidate(graph: &raw::Graph, limits: &super::Limits) -> Result<()> {
    raw::admit_graph(graph, limits)?;
    validate_raw_graph_metadata(graph)?;
    graph.catalog.validate_for_limits(limits)?;
    let plan = plan_write_with_limits(&graph.catalog, graph.conformance, limits)?;
    if plan.bytes > limits.xml_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::XmlBytes,
            actual: plan.bytes,
            maximum: limits.xml_bytes(),
        });
    }
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

fn published_catalog(graph: &raw::Graph, limits: &super::Limits) -> Result<Catalog> {
    if let Some(xml) = &graph.root_xml {
        let (catalog, conformance) = read_with_limits(xml.as_slice(), limits)?;
        if conformance == graph.conformance && catalog == graph.catalog {
            return Ok(graph.catalog.clone());
        }
    }
    let encoded = write_with_limits(&graph.catalog, graph.conformance, limits)?;
    let (catalog, conformance) = super::codec::read_admitted_with_limits(&encoded, limits)?;
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

#[cfg(test)]
mod tests {
    use super::next_relationship_id_from;
    use litchi_opc::{BlobPart, PackURI, Part};

    #[test]
    fn saturated_native_relationship_id_is_bounded() {
        let mut part = BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            "application/xml".to_owned(),
            Vec::new(),
        );
        part.rels_mut().add_relationship(
            "urn:test".to_owned(),
            "native.bin".to_owned(),
            "rId4294967295".to_owned(),
            false,
        );
        let error = next_relationship_id_from(&part, u32::MAX).unwrap_err();
        assert!(error.to_string().contains("no glossary relationship ID"));
    }
}
