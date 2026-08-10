//! Durable advanced-layout transactions for `styles.xml`.

use std::collections::{BTreeMap, BTreeSet};

use base64::{Engine as _, engine::general_purpose::STANDARD as BASE64};
use litchi_core::Result;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::{NsReader, Reader};
use serde_json::{Value, json};

use super::{Kind, Master, bad, read, set_text, set_xml};
use crate::page_layout::{PageLayout, parse_page_layouts, set_page_layout_xml};

const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_CHANGES: usize = 65_536;
const MAX_DURABLE_BYTES: usize = 192 * 1024 * 1024;
const DURABLE_FORMAT: &str = "litchi.odt.layout.v1";

/// One independently mergeable advanced-layout owner.
#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub enum Target {
    /// One named page-layout definition.
    PageLayout(String),
    /// One named master-page definition, including all six header/footer regions.
    MasterPage(String),
}

/// An immutable, fully parsed `styles.xml` layout snapshot.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    source: String,
    masters: Vec<Master>,
    layouts: Vec<PageLayout>,
}

impl Snapshot {
    /// Parse advanced page-layout and master-page owners from exact XML.
    pub fn parse(source: impl Into<String>) -> Result<Self> {
        let source = source.into();
        if source.len() > MAX_XML_BYTES {
            return Err(bad("layout styles XML exceeds its byte limit"));
        }
        let masters = read(&source)?;
        let layouts = parse_page_layouts(&source)?;
        Ok(Self {
            source,
            masters,
            layouts,
        })
    }

    /// Borrow the exact accepted `styles.xml` source.
    pub fn source_xml(&self) -> &str {
        &self.source
    }

    /// Borrow master pages in document order.
    pub fn master_pages(&self) -> &[Master] {
        &self.masters
    }

    /// Borrow page layouts in document order.
    pub fn page_layouts(&self) -> &[PageLayout] {
        &self.layouts
    }

    /// Start a clone-staged failure-atomic layout edit.
    pub fn edit(&self) -> Transaction {
        Transaction {
            before: self.clone(),
            draft: self.source.clone(),
            changes: BTreeMap::new(),
        }
    }

    /// Prepare a bounded master-page transfer with its page-layout dependency.
    pub fn prepare_master_page_transfer(&self, name: &str) -> Result<Transfer> {
        let master = self
            .masters
            .iter()
            .find(|master| master.name == name)
            .ok_or_else(|| bad(format!("master page '{name}' does not exist")))?;
        let layout_name = master
            .page_layout_name
            .as_deref()
            .ok_or_else(|| bad(format!("master page '{name}' has no page-layout reference")))?;
        let layout = self
            .layouts
            .iter()
            .find(|layout| layout.name == layout_name)
            .ok_or_else(|| {
                bad(format!(
                    "master page '{name}' references missing page layout '{layout_name}'"
                ))
            })?;
        let namespaces = root_namespace_declarations(&self.source)?;
        let master_xml = add_namespace_declarations(&master.xml, &namespaces)?;
        let page_layout_xml = add_namespace_declarations(&layout.xml, &namespaces)?;
        let mut dependencies = linked_dependencies(&master_xml)?;
        dependencies.extend(linked_dependencies(&page_layout_xml)?);
        dependencies.sort();
        dependencies.dedup();
        Ok(Transfer {
            master_name: master.name.clone(),
            master_xml,
            page_layout_name: layout.name.clone(),
            page_layout_xml,
            dependencies,
        })
    }
}

/// One inert linked resource referenced by transferred layout XML.
#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub struct Dependency {
    href: String,
}

impl Dependency {
    /// Borrow the retained `xlink:href` value. No resource is fetched.
    pub fn href(&self) -> &str {
        &self.href
    }
}

/// A detached master page and its exact referenced page-layout definition.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Transfer {
    master_name: String,
    master_xml: String,
    page_layout_name: String,
    page_layout_xml: String,
    dependencies: Vec<Dependency>,
}

impl Transfer {
    /// Name of the transferred master page.
    pub fn master_page_name(&self) -> &str {
        &self.master_name
    }

    /// Name of the transferred page layout.
    pub fn page_layout_name(&self) -> &str {
        &self.page_layout_name
    }

    /// Inert linked resources the destination must resolve separately.
    pub fn dependencies(&self) -> &[Dependency] {
        &self.dependencies
    }
}

/// A staged advanced-layout candidate derived from one immutable snapshot.
#[derive(Clone, Debug)]
pub struct Transaction {
    before: Snapshot,
    draft: String,
    changes: BTreeMap<Target, Change>,
}

impl Transaction {
    /// Borrow the exact XML currently staged.
    pub fn source_xml(&self) -> &str {
        &self.draft
    }

    /// Replace one master page with a complete typed value.
    pub fn replace_master_page(&mut self, page: &Master) -> Result<()> {
        let target = Target::MasterPage(page.name.clone());
        let before = owner_xml(&self.draft, &target)?;
        let fragment = page.to_xml_fragment()?;
        let next =
            litchi_odf_common::style::master::writer::replace(&self.draft, &page.name, &fragment)?;
        self.publish(target, before, Some(fragment), next)
    }

    /// Insert one complete typed master page.
    pub fn insert_master_page(&mut self, page: &Master) -> Result<()> {
        let target = Target::MasterPage(page.name.clone());
        let before = owner_xml(&self.draft, &target)?;
        let fragment = page.to_xml_fragment()?;
        let next = litchi_odf_common::style::master::writer::insert(&self.draft, &fragment)?;
        self.publish(target, before, Some(fragment), next)
    }

    /// Remove one named master page.
    pub fn remove_master_page(&mut self, name: &str) -> Result<()> {
        let target = Target::MasterPage(name.to_string());
        let before = owner_xml(&self.draft, &target)?;
        let next = litchi_odf_common::style::master::writer::remove(&self.draft, name)?;
        self.publish(target, before, None, next)
    }

    /// Set plain text in one standard header/footer region.
    pub fn set_region_text(&mut self, master: &str, kind: Kind, text: &str) -> Result<()> {
        self.change_master(master, |source| set_text(source, master, kind, Some(text)))
    }

    /// Replace one standard header/footer region with complete validated XML.
    pub fn set_region_xml(&mut self, master: &str, kind: Kind, xml: &str) -> Result<()> {
        self.change_master(master, |source| set_xml(source, master, kind, xml))
    }

    /// Remove one standard header/footer region.
    pub fn clear_region(&mut self, master: &str, kind: Kind) -> Result<()> {
        self.change_master(master, |source| set_text(source, master, kind, None))
    }

    /// Replace one existing page-layout definition with exact validated XML.
    pub fn replace_page_layout(&mut self, name: &str, xml: &str) -> Result<()> {
        let target = Target::PageLayout(name.to_string());
        let before = owner_xml(&self.draft, &target)?;
        let next = set_page_layout_xml(&self.draft, name, xml)?;
        self.publish(target, before, Some(xml.to_string()), next)
    }

    /// Insert a transferred master and layout when it has no linked dependencies.
    pub fn insert_transfer(&mut self, transfer: &Transfer) -> Result<()> {
        self.insert_transfer_with(transfer, |_| false)
    }

    /// Insert a transferred master and layout after checking every inert link.
    ///
    /// The resolver only authorizes already-staged destination dependencies;
    /// this owner never fetches or executes linked content.
    pub fn insert_transfer_with(
        &mut self,
        transfer: &Transfer,
        mut dependency_is_available: impl FnMut(&Dependency) -> bool,
    ) -> Result<()> {
        if transfer
            .dependencies
            .iter()
            .any(|dependency| !dependency_is_available(dependency))
        {
            return Err(bad("layout transfer has an unresolved linked dependency"));
        }
        let master_target = Target::MasterPage(transfer.master_name.clone());
        if owner_xml(&self.draft, &master_target)?.is_some() {
            return Err(bad(format!(
                "master page '{}' already exists",
                transfer.master_name
            )));
        }
        let layout_target = Target::PageLayout(transfer.page_layout_name.clone());
        match owner_xml(&self.draft, &layout_target)? {
            Some(existing) if existing != transfer.page_layout_xml => {
                return Err(bad(format!(
                    "page layout '{}' already exists with different XML",
                    transfer.page_layout_name
                )));
            },
            Some(_) => {
                let next = litchi_odf_common::style::master::writer::insert(
                    &self.draft,
                    &transfer.master_xml,
                )?;
                self.publish(master_target, None, Some(transfer.master_xml.clone()), next)?;
            },
            None => {
                let with_layout = insert_page_layout(
                    &self.draft,
                    &transfer.page_layout_name,
                    &transfer.page_layout_xml,
                )?;
                self.publish(
                    layout_target,
                    None,
                    Some(transfer.page_layout_xml.clone()),
                    with_layout,
                )?;
                let next = litchi_odf_common::style::master::writer::insert(
                    &self.draft,
                    &transfer.master_xml,
                )?;
                self.publish(master_target, None, Some(transfer.master_xml.clone()), next)?;
            },
        }
        Ok(())
    }

    /// Restore the exact source candidate and discard staged changes.
    pub fn rollback(&mut self) {
        self.draft.clone_from(&self.before.source);
        self.changes.clear();
    }

    /// Parse, validate, and publish the candidate atomically.
    pub fn commit(self) -> Result<Commit> {
        let snapshot = Snapshot::parse(self.draft)?;
        let patch = Patch {
            source: self.before.source,
            target: snapshot.source.clone(),
            changes: self.changes.into_values().collect(),
        };
        validate_patch(&patch)?;
        Ok(Commit { snapshot, patch })
    }

    fn change_master(
        &mut self,
        name: &str,
        edit: impl FnOnce(&str) -> Result<String>,
    ) -> Result<()> {
        let target = Target::MasterPage(name.to_string());
        let before = owner_xml(&self.draft, &target)?;
        let next = edit(&self.draft)?;
        let after = owner_xml(&next, &target)?;
        self.publish(target, before, after, next)
    }

    fn publish(
        &mut self,
        target: Target,
        before: Option<String>,
        after: Option<String>,
        next: String,
    ) -> Result<()> {
        if next.len() > MAX_XML_BYTES {
            return Err(bad("staged layout XML exceeds its byte limit"));
        }
        if let Some(change) = self.changes.get_mut(&target) {
            change.after = after;
            if change.before == change.after {
                self.changes.remove(&target);
            }
        } else if before != after {
            if self.changes.len() >= MAX_CHANGES {
                return Err(bad("layout transaction exceeds its change limit"));
            }
            self.changes.insert(
                target.clone(),
                Change {
                    target,
                    before,
                    after,
                },
            );
        }
        self.draft = next;
        Ok(())
    }
}

/// A validated advanced-layout publication.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the committed immutable snapshot.
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the exact-source reversible patch.
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit and return its immutable snapshot.
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Change {
    target: Target,
    before: Option<String>,
    after: Option<String>,
}

/// An exact-source reversible advanced-layout patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    source: String,
    target: String,
    changes: Vec<Change>,
}

impl Patch {
    /// Whether the patch leaves its exact XML source unchanged.
    pub fn is_empty(&self) -> bool {
        self.source == self.target
    }

    /// Apply only to the exact source XML from which this patch was built.
    pub fn apply(&self, source: &Snapshot) -> Result<Commit> {
        if source.source != self.source {
            return Err(bad("layout patch source snapshot does not match"));
        }
        let snapshot = Snapshot::parse(self.target.clone())?;
        Ok(Commit {
            snapshot,
            patch: self.clone(),
        })
    }

    /// Return the exact-source patch that restores the accepted XML.
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            changes: self
                .changes
                .iter()
                .rev()
                .map(|change| Change {
                    target: change.target.clone(),
                    before: change.after.clone(),
                    after: change.before.clone(),
                })
                .collect(),
        }
    }

    /// Convert this semantic patch into bounded canonical JSON.
    pub fn durable(&self) -> Result<DurablePatch> {
        DurablePatch::new(self.clone())
    }

    /// Build a target-aware three-way merge plan for two patches from one source.
    pub fn merge(left: &Self, right: &Self) -> Result<MergePlan> {
        if left.source != right.source {
            return Err(bad("layout merge branches do not share one source"));
        }
        let source = left.source.clone();
        let left = changes_by_target(&left.changes)?;
        let right = changes_by_target(&right.changes)?;
        let mut selected = BTreeMap::new();
        let mut conflicts = BTreeSet::new();
        for target in left
            .keys()
            .chain(right.keys())
            .cloned()
            .collect::<BTreeSet<_>>()
        {
            match (left.get(&target), right.get(&target)) {
                (Some(left), Some(right)) if left.after != right.after => {
                    conflicts.insert(target);
                },
                (Some(left), _) => {
                    selected.insert(target, left.clone());
                },
                (None, Some(right)) => {
                    selected.insert(target, right.clone());
                },
                (None, None) => {},
            }
        }
        Ok(MergePlan {
            source,
            left,
            right,
            selected,
            conflicts,
        })
    }
}

/// Which branch supplies a conflicting layout owner.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Resolution {
    /// Keep the left branch owner.
    Left,
    /// Keep the right branch owner.
    Right,
}

/// A target-aware advanced-layout merge with explicit conflict resolution.
#[derive(Clone, Debug)]
pub struct MergePlan {
    source: String,
    left: BTreeMap<Target, Change>,
    right: BTreeMap<Target, Change>,
    selected: BTreeMap<Target, Change>,
    conflicts: BTreeSet<Target>,
}

impl MergePlan {
    /// Return unresolved owner conflicts in deterministic order.
    pub fn conflicts(&self) -> impl ExactSizeIterator<Item = &Target> {
        self.conflicts.iter()
    }

    /// Resolve one conflicting owner from a named branch.
    pub fn resolve(&mut self, target: &Target, resolution: Resolution) -> Result<()> {
        if !self.conflicts.remove(target) {
            return Err(bad("layout target is not an unresolved conflict"));
        }
        let branch = match resolution {
            Resolution::Left => &self.left,
            Resolution::Right => &self.right,
        };
        let change = branch
            .get(target)
            .ok_or_else(|| bad("layout merge branch is missing its conflict owner"))?;
        self.selected.insert(target.clone(), change.clone());
        Ok(())
    }

    /// Replay all selected semantic changes after every conflict is resolved.
    pub fn finish(self) -> Result<Patch> {
        if !self.conflicts.is_empty() {
            return Err(bad("layout merge still has unresolved conflicts"));
        }
        let mut target = self.source.clone();
        let changes: Vec<_> = self.selected.into_values().collect();
        for change in &changes {
            target = apply_change(&target, change)?;
        }
        let patch = Patch {
            source: self.source,
            target,
            changes,
        };
        validate_patch(&patch)?;
        Ok(patch)
    }
}

/// Bounded deterministic-JSON advanced-layout patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DurablePatch {
    patch: Patch,
}

impl DurablePatch {
    fn new(patch: Patch) -> Result<Self> {
        validate_patch(&patch)?;
        Ok(Self { patch })
    }

    /// Parse and validate a canonical durable layout patch.
    pub fn from_deterministic_json(bytes: &[u8]) -> Result<Self> {
        if bytes.len() > MAX_DURABLE_BYTES {
            return Err(bad("durable layout patch exceeds its byte limit"));
        }
        let value: Value = serde_json::from_slice(bytes)
            .map_err(|error| bad(format!("invalid durable layout patch: {error}")))?;
        let durable = Self::new(patch_from_value(&value)?)?;
        if durable.to_deterministic_json()? != bytes {
            return Err(bad("durable layout patch is not canonical JSON"));
        }
        Ok(durable)
    }

    /// Serialize this semantic patch as canonical deterministic JSON.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>> {
        let bytes = serde_json::to_vec(&patch_value(&self.patch))
            .map_err(|error| bad(format!("durable layout patch write failed: {error}")))?;
        if bytes.len() > MAX_DURABLE_BYTES {
            return Err(bad("durable layout patch exceeds its byte limit"));
        }
        Ok(bytes)
    }

    /// Apply after checking the complete source layout XML.
    pub fn apply(&self, source: &Snapshot) -> Result<Commit> {
        self.patch.apply(source)
    }

    /// Return the durable patch that restores the exact source XML.
    pub fn inverse(&self) -> Self {
        Self {
            patch: self.patch.inverse(),
        }
    }
}

fn validate_patch(patch: &Patch) -> Result<()> {
    if patch.source.len() > MAX_XML_BYTES
        || patch.target.len() > MAX_XML_BYTES
        || patch.changes.len() > MAX_CHANGES
    {
        return Err(bad("layout patch exceeds its finite bounds"));
    }
    Snapshot::parse(patch.source.clone())?;
    Snapshot::parse(patch.target.clone())?;
    let mut replay = patch.source.clone();
    let mut targets = BTreeSet::new();
    for change in &patch.changes {
        if !targets.insert(change.target.clone()) {
            return Err(bad("layout patch repeats one semantic target"));
        }
        replay = apply_change(&replay, change)?;
    }
    if replay != patch.target {
        return Err(bad("layout patch target is not its semantic replay result"));
    }
    Ok(())
}

fn apply_change(source: &str, change: &Change) -> Result<String> {
    if owner_xml(source, &change.target)? != change.before {
        return Err(bad("layout change semantic precondition does not match"));
    }
    match (&change.target, &change.before, &change.after) {
        (Target::MasterPage(_), None, Some(after)) => {
            litchi_odf_common::style::master::writer::insert(source, after)
        },
        (Target::MasterPage(_), Some(before), Some(after)) => {
            replace_owner_exact(source, before, after, &change.target)
        },
        (Target::MasterPage(name), Some(_), None) => {
            litchi_odf_common::style::master::writer::remove(source, name)
        },
        (Target::PageLayout(name), None, Some(after)) => insert_page_layout(source, name, after),
        (Target::PageLayout(_), Some(before), Some(after)) => {
            replace_owner_exact(source, before, after, &change.target)
        },
        (Target::PageLayout(name), Some(before), None) => remove_page_layout(source, name, before),
        _ => Err(bad("invalid layout semantic change")),
    }
}

fn replace_owner_exact(source: &str, before: &str, after: &str, target: &Target) -> Result<String> {
    let start = source
        .find(before)
        .ok_or_else(|| bad("layout owner XML span is missing"))?;
    let candidate = super::replace_range(source, start, start + before.len(), after)?;
    if owner_xml(&candidate, target)?.as_deref() != Some(after) {
        return Err(bad("replaced layout owner did not validate exactly"));
    }
    Ok(candidate)
}

fn owner_xml(source: &str, target: &Target) -> Result<Option<String>> {
    match target {
        Target::MasterPage(name) => Ok(read(source)?
            .into_iter()
            .find(|master| master.name == *name)
            .map(|master| master.xml)),
        Target::PageLayout(name) => Ok(parse_page_layouts(source)?
            .into_iter()
            .find(|layout| layout.name == *name)
            .map(|layout| layout.xml)),
    }
}

fn insert_page_layout(source: &str, name: &str, fragment: &str) -> Result<String> {
    if owner_xml(source, &Target::PageLayout(name.to_string()))?.is_some() {
        return Err(bad(format!("page layout '{name}' already exists")));
    }
    let masters = read(source)?;
    let mut suffix = 0usize;
    let temporary = loop {
        let candidate = format!("__litchi_layout_transfer_{suffix}");
        if masters.iter().all(|master| master.name != candidate) {
            break candidate;
        }
        suffix = suffix
            .checked_add(1)
            .ok_or_else(|| bad("temporary layout owner name overflow"))?;
    };
    let source = litchi_odf_common::style::master::writer::add(source, &temporary, name)?;
    let source = set_page_layout_xml(&source, name, fragment)?;
    litchi_odf_common::style::master::writer::remove(&source, &temporary)
}

fn remove_page_layout(source: &str, name: &str, expected: &str) -> Result<String> {
    let actual = owner_xml(source, &Target::PageLayout(name.to_string()))?
        .ok_or_else(|| bad(format!("page layout '{name}' does not exist")))?;
    if actual != expected {
        return Err(bad("page-layout removal precondition does not match"));
    }
    let start = source
        .find(&actual)
        .ok_or_else(|| bad("page-layout XML span is missing"))?;
    super::replace_range(source, start, start + actual.len(), "")
}

fn changes_by_target(changes: &[Change]) -> Result<BTreeMap<Target, Change>> {
    let mut result = BTreeMap::new();
    for change in changes {
        if result
            .insert(change.target.clone(), change.clone())
            .is_some()
        {
            return Err(bad("layout patch repeats one semantic target"));
        }
    }
    Ok(result)
}

fn linked_dependencies(xml: &str) -> Result<Vec<Dependency>> {
    const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut dependencies = Vec::new();
    let mut buffer = Vec::new();
    loop {
        let (_, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| bad(format!("invalid transferred layout XML: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes() {
                    let attribute = attribute
                        .map_err(|error| bad(format!("invalid layout attribute: {error}")))?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    if matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == XLINK)
                        && local.as_ref() == b"href"
                    {
                        let href = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                            .map_err(|error| bad(format!("invalid layout link: {error}")))?
                            .into_owned();
                        if !href.is_empty() {
                            dependencies.push(Dependency { href });
                        }
                    }
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(bad("active XML is forbidden in layout transfers"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(dependencies)
}

fn root_namespace_declarations(source: &str) -> Result<Vec<(String, String)>> {
    let mut reader = Reader::from_str(source);
    loop {
        match reader
            .read_event()
            .map_err(|error| bad(format!("invalid layout XML: {error}")))?
        {
            Event::Start(element) | Event::Empty(element) => {
                let mut namespaces = Vec::new();
                for attribute in element.attributes() {
                    let attribute = attribute
                        .map_err(|error| bad(format!("invalid namespace attribute: {error}")))?;
                    let name = std::str::from_utf8(attribute.key.as_ref())
                        .map_err(|error| bad(format!("invalid namespace name: {error}")))?;
                    if name == "xmlns" || name.starts_with("xmlns:") {
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                            .map_err(|error| bad(format!("invalid namespace value: {error}")))?
                            .into_owned();
                        namespaces.push((name.to_string(), value));
                    }
                }
                return Ok(namespaces);
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(bad("active XML is forbidden in layout styles"));
            },
            Event::Eof => return Err(bad("layout styles XML has no document element")),
            _ => {},
        }
    }
}

fn add_namespace_declarations(fragment: &str, namespaces: &[(String, String)]) -> Result<String> {
    let mut reader = Reader::from_str(fragment);
    let element = loop {
        match reader
            .read_event()
            .map_err(|error| bad(format!("invalid layout transfer fragment: {error}")))?
        {
            Event::Start(element) | Event::Empty(element) => break element.into_owned(),
            Event::DocType(_) | Event::PI(_) => {
                return Err(bad("active XML is forbidden in layout transfer fragments"));
            },
            Event::Eof => return Err(bad("layout transfer fragment has no root element")),
            _ => {},
        }
    };
    let declared: BTreeSet<String> = element
        .attributes()
        .filter_map(|attribute| attribute.ok())
        .filter_map(|attribute| {
            let name = std::str::from_utf8(attribute.key.as_ref()).ok()?;
            (name == "xmlns" || name.starts_with("xmlns:")).then(|| name.to_string())
        })
        .collect();
    let missing: Vec<_> = namespaces
        .iter()
        .filter(|(name, _)| !declared.contains(name))
        .collect();
    if missing.is_empty() {
        return Ok(fragment.to_string());
    }
    let close = first_start_tag_close(fragment)?;
    let insertion = if fragment.as_bytes().get(close.wrapping_sub(1)) == Some(&b'/') {
        close - 1
    } else {
        close
    };
    let mut output = String::with_capacity(
        fragment.len()
            + missing
                .iter()
                .map(|(name, value)| name.len() + value.len() + 4)
                .sum::<usize>(),
    );
    output.push_str(&fragment[..insertion]);
    for (name, value) in missing {
        output.push(' ');
        output.push_str(name);
        output.push_str("=\"");
        output.push_str(&litchi_core::xml::escape_xml(value));
        output.push('"');
    }
    output.push_str(&fragment[insertion..]);
    Ok(output)
}

fn first_start_tag_close(xml: &str) -> Result<usize> {
    let mut quote = None;
    let mut started = false;
    for (index, byte) in xml.bytes().enumerate() {
        if !started {
            if byte == b'<' {
                started = true;
            }
            continue;
        }
        match (quote, byte) {
            (None, b'\'' | b'"') => quote = Some(byte),
            (Some(active), value) if active == value => quote = None,
            (None, b'>') => return Ok(index),
            _ => {},
        }
    }
    Err(bad("layout transfer root start tag is unterminated"))
}

fn patch_value(patch: &Patch) -> Value {
    json!({
        "changes": patch.changes.iter().map(change_value).collect::<Vec<_>>(),
        "format": DURABLE_FORMAT,
        "source": BASE64.encode(patch.source.as_bytes()),
        "target": BASE64.encode(patch.target.as_bytes()),
    })
}

fn change_value(change: &Change) -> Value {
    let (kind, name) = match &change.target {
        Target::PageLayout(name) => ("page-layout", name),
        Target::MasterPage(name) => ("master-page", name),
    };
    json!({
        "after": change.after.as_deref(),
        "before": change.before.as_deref(),
        "kind": kind,
        "name": name,
    })
}

fn patch_from_value(value: &Value) -> Result<Patch> {
    let object = value
        .as_object()
        .ok_or_else(|| bad("durable layout patch must be an object"))?;
    if object.len() != 4 || object.get("format").and_then(Value::as_str) != Some(DURABLE_FORMAT) {
        return Err(bad("unknown durable layout patch envelope"));
    }
    let changes = object
        .get("changes")
        .and_then(Value::as_array)
        .ok_or_else(|| bad("durable layout patch changes are missing"))?;
    if changes.len() > MAX_CHANGES {
        return Err(bad("durable layout patch exceeds its change limit"));
    }
    Ok(Patch {
        source: decode_xml(object.get("source"), "source")?,
        target: decode_xml(object.get("target"), "target")?,
        changes: changes
            .iter()
            .map(change_from_value)
            .collect::<Result<_>>()?,
    })
}

fn change_from_value(value: &Value) -> Result<Change> {
    let object = value
        .as_object()
        .ok_or_else(|| bad("durable layout change must be an object"))?;
    if object.len() != 4 {
        return Err(bad("durable layout change has unknown fields"));
    }
    let name = object
        .get("name")
        .and_then(Value::as_str)
        .ok_or_else(|| bad("durable layout change name is missing"))?;
    if name.is_empty() || name.len() > super::MAX_VALUE {
        return Err(bad("invalid durable layout owner name"));
    }
    let target = match object.get("kind").and_then(Value::as_str) {
        Some("page-layout") => Target::PageLayout(name.to_string()),
        Some("master-page") => Target::MasterPage(name.to_string()),
        _ => return Err(bad("unknown durable layout owner kind")),
    };
    Ok(Change {
        target,
        before: optional_xml(object.get("before"), "before")?,
        after: optional_xml(object.get("after"), "after")?,
    })
}

fn decode_xml(value: Option<&Value>, name: &str) -> Result<String> {
    let value = value
        .and_then(Value::as_str)
        .ok_or_else(|| bad(format!("durable layout patch {name} is missing")))?;
    let bytes = BASE64
        .decode(value)
        .map_err(|error| bad(format!("invalid durable layout patch {name}: {error}")))?;
    if bytes.len() > MAX_XML_BYTES {
        return Err(bad("durable layout XML exceeds its byte limit"));
    }
    String::from_utf8(bytes)
        .map_err(|error| bad(format!("durable layout patch {name} is not UTF-8: {error}")))
}

fn optional_xml(value: Option<&Value>, name: &str) -> Result<Option<String>> {
    match value {
        Some(Value::Null) => Ok(None),
        Some(Value::String(value)) if value.len() <= MAX_XML_BYTES => Ok(Some(value.clone())),
        _ => Err(bad(format!("invalid durable layout change {name}"))),
    }
}
