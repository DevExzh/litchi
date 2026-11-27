//! Durable advanced-layout transactions for `styles.xml`.

use std::collections::{BTreeMap, BTreeSet};

use litchi_core::Result;
use litchi_odf_common::XmlStreamLimits;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::{NsReader, Reader};

use super::{Master, bad, read};
use crate::page_layout::{PageLayout, parse_page_layouts, set_page_layout_xml};
use crate::section_properties::{SectionStyleProperties, parse_section_style_properties};
use crate::style::columns::Columns;

const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_CHANGES: usize = 65_536;

/// One independently mergeable advanced-layout owner.
#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub(crate) enum Target {
    /// One named page-layout definition.
    PageLayout(String),
    /// One named section-layout style, including columns and background.
    SectionStyle(String),
    /// One named master-page definition, including all six header/footer regions.
    MasterPage(String),
}

/// An immutable, fully parsed `styles.xml` layout snapshot.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    source: String,
    masters: Vec<Master>,
    layouts: Vec<PageLayout>,
    sections: Vec<SectionLayout>,
    limits: XmlStreamLimits,
}

impl Snapshot {
    /// Parse advanced page-layout, section-style, and master-page owners.
    pub fn parse(source: impl Into<String>, limits: XmlStreamLimits) -> Result<Self> {
        let source = source.into();
        if source.len() > MAX_XML_BYTES {
            return Err(bad("layout styles XML exceeds its byte limit"));
        }
        let masters = read(&source, limits).map_err(litchi_core::Error::from)?;
        let layouts = parse_page_layouts(&source)?;
        let sections = section_layouts(&source, &limits)?;
        Ok(Self {
            source,
            masters,
            layouts,
            sections,
            limits,
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

    /// Borrow typed section-layout styles in source order.
    pub fn section_layouts(&self) -> &[SectionLayout] {
        &self.sections
    }

    /// Start a clone-staged failure-atomic layout edit.
    pub(crate) fn edit(&self) -> Transaction {
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

    /// Prepare one exact section style with typed columns and inert link dependencies.
    pub fn prepare_section_layout_transfer(&self, name: &str) -> Result<SectionTransfer> {
        let section = self
            .sections
            .iter()
            .find(|section| section.name() == name)
            .ok_or_else(|| bad(format!("section style '{name}' does not exist")))?;
        let namespaces = root_namespace_declarations(&self.source)?;
        let xml = add_namespace_declarations(&section.xml, &namespaces)?;
        let mut dependencies = linked_dependencies(&xml)?;
        dependencies.sort();
        dependencies.dedup();
        Ok(SectionTransfer {
            name: name.to_string(),
            xml,
            dependencies,
        })
    }
}

/// One typed section style with its exact source-backed XML owner.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SectionLayout {
    style: SectionStyleProperties,
    columns: Option<Columns>,
    xml: String,
}

impl SectionLayout {
    /// Create a compact authored section layout from typed properties and columns.
    pub fn new(style: SectionStyleProperties, columns: Option<Columns>) -> Result<Self> {
        style.validate()?;
        if let Some(columns) = &columns {
            columns.validate()?;
        }
        let xml = section_style_fragment(&style, columns.as_ref())?;
        Ok(Self {
            style,
            columns,
            xml,
        })
    }

    /// Exact section-style name.
    pub fn name(&self) -> &str {
        &self.style.name
    }

    /// Typed margins, background, writing mode, and protection metadata.
    pub fn properties(&self) -> &SectionStyleProperties {
        &self.style
    }

    /// Typed multi-column section layout, when declared.
    pub fn columns(&self) -> Option<&Columns> {
        self.columns.as_ref()
    }

    /// Exact source-backed or compact authored style XML.
    pub fn xml(&self) -> &str {
        &self.xml
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

/// A detached exact section-layout owner with inert dependency metadata.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SectionTransfer {
    name: String,
    xml: String,
    dependencies: Vec<Dependency>,
}

impl SectionTransfer {
    /// Name of the transferred section style.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Inert linked resources the destination must resolve separately.
    pub fn dependencies(&self) -> &[Dependency] {
        &self.dependencies
    }

    /// Exact self-contained XML retained for destination insertion.
    pub fn xml(&self) -> &str {
        &self.xml
    }
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
pub(crate) struct Transaction {
    before: Snapshot,
    draft: String,
    changes: BTreeMap<Target, Change>,
}

impl Transaction {
    /// Insert one compact typed section-layout style.
    pub(crate) fn insert_section_layout(&mut self, section: &SectionLayout) -> Result<()> {
        let target = Target::SectionStyle(section.name().to_string());
        let before = owner_xml(&self.draft, &target, self.before.limits)?;
        if before.is_some() {
            return Err(bad(format!(
                "section style '{}' already exists",
                section.name()
            )));
        }
        let next = insert_section_style(&self.draft, section.xml(), self.before.limits)?;
        self.publish(target, before, Some(section.xml().to_string()), next)
    }

    /// Replace one named section-layout style with a typed compact owner.
    pub(crate) fn replace_section_layout(&mut self, section: &SectionLayout) -> Result<()> {
        let target = Target::SectionStyle(section.name().to_string());
        let before = owner_xml(&self.draft, &target, self.before.limits)?;
        let current = before
            .as_deref()
            .ok_or_else(|| bad(format!("section style '{}' does not exist", section.name())))?;
        let next = replace_owner_exact(
            &self.draft,
            current,
            section.xml(),
            &target,
            self.before.limits,
        )?;
        self.publish(target, before, Some(section.xml().to_string()), next)
    }

    /// Remove one named section-layout style.
    pub(crate) fn remove_section_layout(&mut self, name: &str) -> Result<()> {
        let target = Target::SectionStyle(name.to_string());
        let before = owner_xml(&self.draft, &target, self.before.limits)?;
        let current = before
            .as_deref()
            .ok_or_else(|| bad(format!("section style '{name}' does not exist")))?;
        let next = remove_owner_exact(&self.draft, current, &target, self.before.limits)?;
        self.publish(target, before, None, next)
    }

    #[cfg(test)]
    /// Insert a transferred master and layout when it has no linked dependencies.
    pub(crate) fn insert_transfer(&mut self, transfer: &Transfer) -> Result<()> {
        self.insert_transfer_with(transfer, |_| false)
    }

    /// Insert a transferred master and layout after checking every inert link.
    ///
    /// The resolver only authorizes already-staged destination dependencies;
    /// this owner never fetches or executes linked content.
    pub(crate) fn insert_transfer_with(
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
        if owner_xml(&self.draft, &master_target, self.before.limits)?.is_some() {
            return Err(bad(format!(
                "master page '{}' already exists",
                transfer.master_name
            )));
        }
        let layout_target = Target::PageLayout(transfer.page_layout_name.clone());
        match owner_xml(&self.draft, &layout_target, self.before.limits)? {
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
                    self.before.limits,
                )?;
                self.publish(master_target, None, Some(transfer.master_xml.clone()), next)?;
            },
            None => {
                let with_layout = insert_page_layout(
                    &self.draft,
                    &transfer.page_layout_name,
                    &transfer.page_layout_xml,
                    self.before.limits,
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
                    self.before.limits,
                )?;
                self.publish(master_target, None, Some(transfer.master_xml.clone()), next)?;
            },
        }
        Ok(())
    }

    #[cfg(test)]
    /// Insert a transferred section layout when it has no linked dependencies.
    pub(crate) fn insert_section_transfer(&mut self, transfer: &SectionTransfer) -> Result<()> {
        self.insert_section_transfer_with(transfer, |_| false)
    }

    /// Insert a transferred section layout after checking every inert link.
    pub(crate) fn insert_section_transfer_with(
        &mut self,
        transfer: &SectionTransfer,
        mut dependency_is_available: impl FnMut(&Dependency) -> bool,
    ) -> Result<()> {
        if transfer
            .dependencies
            .iter()
            .any(|dependency| !dependency_is_available(dependency))
        {
            return Err(bad(
                "section-layout transfer has an unresolved linked dependency",
            ));
        }
        let target = Target::SectionStyle(transfer.name.clone());
        if owner_xml(&self.draft, &target, self.before.limits)?.is_some() {
            return Err(bad(format!(
                "section style '{}' already exists",
                transfer.name
            )));
        }
        let next = insert_section_style(&self.draft, &transfer.xml, self.before.limits)?;
        self.publish(target, None, Some(transfer.xml.clone()), next)
    }

    /// Parse, validate, and publish the candidate atomically.
    pub(crate) fn commit(self) -> Result<Commit> {
        let snapshot = Snapshot::parse(self.draft, self.before.limits)?;
        let patch = Patch {
            source: self.before.source,
            target: snapshot.source.clone(),
            changes: self.changes.into_values().collect(),
        };
        validate_patch(&patch, self.before.limits)?;
        Ok(Commit { snapshot })
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
pub(crate) struct Commit {
    snapshot: Snapshot,
}

impl Commit {
    #[cfg(test)]
    /// Borrow the committed immutable snapshot.
    pub(crate) fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Consume the commit and return its immutable snapshot.
    pub(crate) fn into_snapshot(self) -> Snapshot {
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
pub(crate) struct Patch {
    source: String,
    target: String,
    changes: Vec<Change>,
}

fn validate_patch(patch: &Patch, limits: XmlStreamLimits) -> Result<()> {
    if patch.source.len() > MAX_XML_BYTES
        || patch.target.len() > MAX_XML_BYTES
        || patch.changes.len() > MAX_CHANGES
    {
        return Err(bad("layout patch exceeds its finite bounds"));
    }
    Snapshot::parse(patch.source.clone(), limits)?;
    Snapshot::parse(patch.target.clone(), limits)?;
    let mut replay = patch.source.clone();
    let mut targets = BTreeSet::new();
    for change in &patch.changes {
        if !targets.insert(change.target.clone()) {
            return Err(bad("layout patch repeats one semantic target"));
        }
        replay = apply_change(&replay, change, limits)?;
    }
    if replay != patch.target {
        return Err(bad("layout patch target is not its semantic replay result"));
    }
    Ok(())
}

fn apply_change(source: &str, change: &Change, limits: XmlStreamLimits) -> Result<String> {
    if owner_xml(source, &change.target, limits)? != change.before {
        return Err(bad("layout change semantic precondition does not match"));
    }
    match (&change.target, &change.before, &change.after) {
        (Target::MasterPage(_), None, Some(after)) => {
            litchi_odf_common::style::master::writer::insert(source, after, limits)
                .map_err(litchi_core::Error::from)
        },
        (Target::MasterPage(_), Some(before), Some(after)) => {
            replace_owner_exact(source, before, after, &change.target, limits)
        },
        (Target::MasterPage(name), Some(_), None) => {
            litchi_odf_common::style::master::writer::remove(source, name, limits)
                .map_err(litchi_core::Error::from)
        },
        (Target::PageLayout(name), None, Some(after)) => {
            insert_page_layout(source, name, after, limits)
        },
        (Target::PageLayout(_), Some(before), Some(after)) => {
            replace_owner_exact(source, before, after, &change.target, limits)
        },
        (Target::PageLayout(name), Some(before), None) => {
            remove_page_layout(source, name, before, limits)
        },
        (Target::SectionStyle(_), None, Some(after)) => insert_section_style(source, after, limits),
        (Target::SectionStyle(_), Some(before), Some(after)) => {
            replace_owner_exact(source, before, after, &change.target, limits)
        },
        (Target::SectionStyle(_), Some(before), None) => {
            remove_owner_exact(source, before, &change.target, limits)
        },
        _ => Err(bad("invalid layout semantic change")),
    }
}

fn replace_owner_exact(
    source: &str,
    before: &str,
    after: &str,
    target: &Target,
    limits: XmlStreamLimits,
) -> Result<String> {
    let start = source
        .find(before)
        .ok_or_else(|| bad("layout owner XML span is missing"))?;
    let candidate = super::replace_range(source, start, start + before.len(), after)?;
    if owner_xml(&candidate, target, limits)?.as_deref() != Some(after) {
        return Err(bad("replaced layout owner did not validate exactly"));
    }
    Ok(candidate)
}

fn remove_owner_exact(
    source: &str,
    before: &str,
    target: &Target,
    limits: XmlStreamLimits,
) -> Result<String> {
    let start = source
        .find(before)
        .ok_or_else(|| bad("layout owner XML span is missing"))?;
    let candidate = super::replace_range(source, start, start + before.len(), "")?;
    if owner_xml(&candidate, target, limits)?.is_some() {
        return Err(bad("removed layout owner remains present"));
    }
    Ok(candidate)
}

fn owner_xml(source: &str, target: &Target, limits: XmlStreamLimits) -> Result<Option<String>> {
    match target {
        Target::MasterPage(name) => Ok(read(source, limits)?
            .into_iter()
            .find(|master| master.name == *name)
            .map(|master| master.xml)),
        Target::PageLayout(name) => Ok(parse_page_layouts(source)?
            .into_iter()
            .find(|layout| layout.name == *name)
            .map(|layout| layout.xml)),
        Target::SectionStyle(name) => Ok(section_style_owners(source)?
            .into_iter()
            .find(|(owner_name, _)| owner_name == name)
            .map(|(_, xml)| xml)),
    }
}

fn insert_page_layout(
    source: &str,
    name: &str,
    fragment: &str,
    limits: XmlStreamLimits,
) -> Result<String> {
    if owner_xml(source, &Target::PageLayout(name.to_string()), limits)?.is_some() {
        return Err(bad(format!("page layout '{name}' already exists")));
    }
    let masters = read(source, limits)?;
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
    let source = litchi_odf_common::style::master::writer::add(source, &temporary, name, limits)?;
    let source = set_page_layout_xml(&source, name, fragment)?;
    litchi_odf_common::style::master::writer::remove(&source, &temporary, limits)
        .map_err(litchi_core::Error::from)
}

fn remove_page_layout(
    source: &str,
    name: &str,
    expected: &str,
    limits: XmlStreamLimits,
) -> Result<String> {
    let actual = owner_xml(source, &Target::PageLayout(name.to_string()), limits)?
        .ok_or_else(|| bad(format!("page layout '{name}' does not exist")))?;
    if actual != expected {
        return Err(bad("page-layout removal precondition does not match"));
    }
    let start = source
        .find(&actual)
        .ok_or_else(|| bad("page-layout XML span is missing"))?;
    super::replace_range(source, start, start + actual.len(), "")
}

fn section_layouts(source: &str, limits: &XmlStreamLimits) -> Result<Vec<SectionLayout>> {
    let typed = parse_section_style_properties(source.as_bytes(), limits)?;
    let namespaces = root_namespace_declarations(source)?;
    let owners = section_style_owners(source)?;
    let mut names = BTreeSet::new();
    let mut result = Vec::new();
    result
        .try_reserve_exact(typed.styles.len())
        .map_err(|error| bad(format!("section layout allocation failed: {error}")))?;
    for (name, xml) in owners {
        if !names.insert(name.clone()) {
            return Err(bad(format!("duplicate section style '{name}'")));
        }
        let Some(style) = typed.get(&name).cloned() else {
            continue;
        };
        let closed = add_namespace_declarations(&xml, &namespaces)?;
        let wrapper = format!(
            "<office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"><office:styles>{closed}</office:styles></office:document-styles>"
        );
        let mut columns = crate::style::columns::parse(&wrapper)?;
        if columns.len() > 1 {
            return Err(bad(format!(
                "section style '{name}' contains multiple column owners"
            )));
        }
        result.push(SectionLayout {
            style,
            columns: columns.pop(),
            xml,
        });
    }
    if result.len() != typed.styles.len() {
        return Err(bad(
            "typed section styles do not match their exact source owners",
        ));
    }
    Ok(result)
}

fn section_style_owners(source: &str) -> Result<Vec<(String, String)>> {
    const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";

    let mut reader = NsReader::from_str(source);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut version = XmlVersion::Implicit1_0;
    let mut active: Option<(String, usize, usize)> = None;
    let mut result = Vec::new();
    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|error| bad(format!("section style position is out of range: {error}")))?;
        let (is_style, event) = {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| bad(format!("invalid section style XML: {error}")))?;
            let is_style =
                matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == STYLE);
            (is_style, event.into_owned())
        };
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|error| bad(format!("section style position is out of range: {error}")))?;
        match event {
            Event::Start(element) => {
                if let Some((_, _, depth)) = active.as_mut() {
                    *depth = depth
                        .checked_add(1)
                        .ok_or_else(|| bad("section style nesting overflow"))?;
                } else if is_style && element.local_name().as_ref() == b"style" {
                    let (family, name) = section_style_identity(&reader, version, &element)?;
                    if family.as_deref() == Some("section") {
                        let name = name.ok_or_else(|| bad("section style lacks style:name"))?;
                        active = Some((name, event_start, 1));
                    }
                }
            },
            Event::Empty(element)
                if active.is_none() && is_style && element.local_name().as_ref() == b"style" =>
            {
                let (family, name) = section_style_identity(&reader, version, &element)?;
                if family.as_deref() == Some("section") {
                    let name = name.ok_or_else(|| bad("section style lacks style:name"))?;
                    result.push((name, source[event_start..event_end].to_string()));
                }
            },
            Event::End(_) if active.is_some() => {
                let complete = {
                    let (_, _, depth) = active
                        .as_mut()
                        .ok_or_else(|| bad("missing active section style"))?;
                    *depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| bad("section style nesting underflow"))?;
                    *depth == 0
                };
                if complete {
                    let (name, start, _) = active
                        .take()
                        .ok_or_else(|| bad("missing completed section style"))?;
                    result.push((name, source[start..event_end].to_string()));
                }
            },
            Event::Decl(declaration) => {
                version = declaration
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(bad("active XML is forbidden in section layout owners"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if active.is_some() {
        return Err(bad("unterminated section style owner"));
    }
    Ok(result)
}

fn section_style_identity(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<(Option<String>, Option<String>)> {
    const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";

    let mut family = None;
    let mut name = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| bad(format!("invalid section style attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == STYLE) {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| bad(format!("invalid section style attribute: {error}")))?
            .into_owned();
        match local.as_ref() {
            b"family" => family = Some(value),
            b"name" => name = Some(value),
            _ => {},
        }
    }
    Ok((family, name))
}

fn section_style_fragment(
    style: &SectionStyleProperties,
    columns: Option<&Columns>,
) -> Result<String> {
    let mut xml = style.to_xml_fragment()?;
    let Some(columns) = columns else {
        return Ok(xml);
    };
    let fragment = columns.to_xml_fragment()?;
    let properties = xml
        .find("<style:section-properties")
        .ok_or_else(|| bad("authored section style lacks section properties"))?;
    let start_close = xml[properties..]
        .find('>')
        .map(|offset| properties + offset)
        .ok_or_else(|| bad("authored section-properties start tag is unterminated"))?;
    if xml.as_bytes().get(start_close.wrapping_sub(1)) == Some(&b'/') {
        xml.replace_range(start_close - 1..=start_close, ">");
        let close = xml
            .find("</style:style>")
            .ok_or_else(|| bad("authored section style end tag is missing"))?;
        xml.insert_str(close, &format!("{fragment}</style:section-properties>"));
    } else {
        let close = xml
            .find("</style:section-properties>")
            .ok_or_else(|| bad("authored section-properties end tag is missing"))?;
        xml.insert_str(close, &fragment);
    }
    Ok(xml)
}

fn insert_section_style(source: &str, fragment: &str, limits: XmlStreamLimits) -> Result<String> {
    let closed = fragment.to_string();
    let wrapper = format!(
        "<office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"><office:styles>{closed}</office:styles></office:document-styles>"
    );
    let parsed = section_layouts(&wrapper, &limits)?;
    if parsed.len() != 1 || parsed[0].xml != closed {
        return Err(bad(
            "section layout fragment is not exactly one section style",
        ));
    }
    if owner_xml(
        source,
        &Target::SectionStyle(parsed[0].name().to_string()),
        limits,
    )?
    .is_some()
    {
        return Err(bad(format!(
            "section style '{}' already exists",
            parsed[0].name()
        )));
    }
    insert_office_style_child(source, &closed)
}

fn insert_office_style_child(source: &str, fragment: &str) -> Result<String> {
    const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";

    let mut reader = NsReader::from_str(source);
    let mut buffer = Vec::new();
    let mut active: Option<(usize, bool, usize)> = None;
    let mut fallback = None;
    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|error| bad(format!("style container position is out of range: {error}")))?;
        let (is_office, event) = {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| bad(format!("invalid style container XML: {error}")))?;
            let is_office =
                matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == OFFICE);
            (is_office, event.into_owned())
        };
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|error| bad(format!("style container position is out of range: {error}")))?;
        match event {
            Event::Start(element)
                if active.is_none()
                    && is_office
                    && matches!(
                        element.local_name().as_ref(),
                        b"styles" | b"automatic-styles"
                    ) =>
            {
                let preferred = element.local_name().as_ref() == b"styles";
                active = Some((1, preferred, event_end));
            },
            Event::Empty(element)
                if active.is_none()
                    && is_office
                    && matches!(
                        element.local_name().as_ref(),
                        b"styles" | b"automatic-styles"
                    ) =>
            {
                let qualified_name = element.name();
                let name = std::str::from_utf8(qualified_name.as_ref())
                    .map_err(|error| bad(format!("invalid style container name: {error}")))?;
                let mut replacement = source[event_start..event_end].to_string();
                let close = replacement
                    .rfind("/>")
                    .ok_or_else(|| bad("empty style container is malformed"))?;
                replacement.replace_range(close.., &format!(">{fragment}</{name}>"));
                if element.local_name().as_ref() == b"styles" {
                    return super::replace_range(source, event_start, event_end, &replacement);
                }
                fallback = Some((event_start, Some(replacement)));
            },
            Event::Start(_) if active.is_some() => {
                let (depth, _, _) = active
                    .as_mut()
                    .ok_or_else(|| bad("missing active style container"))?;
                *depth = depth
                    .checked_add(1)
                    .ok_or_else(|| bad("style container nesting overflow"))?;
            },
            Event::End(_) if active.is_some() => {
                let (depth, _, _) = active
                    .as_mut()
                    .ok_or_else(|| bad("missing active style container"))?;
                *depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| bad("style container nesting underflow"))?;
                if *depth == 0 {
                    let (_, preferred, insert_at) = active
                        .take()
                        .ok_or_else(|| bad("missing completed style container"))?;
                    if preferred {
                        return super::replace_range(source, event_start, event_start, fragment);
                    }
                    fallback = Some((insert_at, None::<String>));
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    match fallback {
        Some((start, Some(replacement))) => {
            let end = start
                .checked_add(
                    source[start..]
                        .find("/>")
                        .ok_or_else(|| bad("empty automatic-styles container end is missing"))?
                        + 2,
                )
                .ok_or_else(|| bad("style container end overflow"))?;
            super::replace_range(source, start, end, &replacement)
        },
        Some((end, None)) => super::replace_range(source, end, end, fragment),
        None => Err(bad("document has no style container")),
    }
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
