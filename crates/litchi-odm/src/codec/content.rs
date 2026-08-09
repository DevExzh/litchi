//! Master-document content validation.

use litchi_core::Position;
use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{
    borrow::Cow,
    collections::{HashMap, HashSet},
    ops::Range,
};

const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_IDENTITIES: usize = 1_000_000;
const MAX_REFERENCES: usize = 1_000_000;
const MAX_REFERENCE_BYTES: usize = 16 * 1024;
const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XML: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Text,
    Other,
}

#[derive(Clone, Debug)]
struct ActiveSection {
    depth: usize,
    position: Position,
    name: String,
    child_seen: bool,
    source_seen: bool,
}

/// Bounded semantic projection retained by the package snapshot.
#[derive(Clone, Debug, Default)]
pub(crate) struct Semantics {
    references: Vec<crate::model::subdocument::Reference>,
    href_spans: Vec<Range<usize>>,
    tree: crate::model::section::Tree,
    structure: crate::structure::Structure,
    local_section_references: Vec<(String, Range<usize>)>,
}

impl Semantics {
    pub(crate) fn references(&self) -> &[crate::model::subdocument::Reference] {
        &self.references
    }

    pub(crate) fn href_span(&self, reference: usize) -> Option<&Range<usize>> {
        self.href_spans.get(reference)
    }

    pub(crate) const fn tree(&self) -> &crate::model::section::Tree {
        &self.tree
    }

    pub(crate) fn local_section_references(&self) -> &[(String, Range<usize>)] {
        &self.local_section_references
    }

    pub(crate) const fn structure(&self) -> &crate::structure::Structure {
        &self.structure
    }
}

/// Validate a UTF-8 content part before authoring it into a package.
pub(crate) fn validate(xml: &str) -> Result<()> {
    parse(xml).map(|_| ())
}

/// Validates content and projects ordered, inert subdocument references.
pub(crate) fn parse(xml: &str) -> Result<Semantics> {
    if xml.len() > MAX_CONTENT_BYTES {
        return Err(Error::InvalidFormat(
            "content.xml exceeds the family limit".to_string(),
        ));
    }
    parse_structure(xml)
}

fn parse_structure(xml: &str) -> Result<Semantics> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut body_seen = false;
    let mut master_seen = false;
    let mut body_depth = None;
    let mut master_depth = None;
    let mut section_names = HashSet::new();
    let mut xml_ids = HashSet::new();
    let mut sections: Vec<ActiveSection> = Vec::new();
    let mut tree = crate::model::section::Tree::default();
    let mut references = Vec::new();
    let mut href_spans = Vec::new();
    let mut section_source_depth = None;
    let mut local_section_references = Vec::new();
    loop {
        let event_start = position(&reader)?;
        let (resolved_namespace, borrowed_event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODM content XML: {error}")))?;
        let namespace = classify(&resolved_namespace);
        let event = borrowed_event.into_owned();
        let event_end = position(&reader)?;
        match event {
            Event::Start(element) => {
                depth = checked_depth(depth)?;
                observe(
                    &reader,
                    namespace,
                    &element,
                    depth,
                    false,
                    &mut root_seen,
                    &mut body_seen,
                    &mut master_seen,
                    &mut body_depth,
                    &mut master_depth,
                    &mut section_names,
                    &mut xml_ids,
                    &mut sections,
                    &mut tree,
                    &mut references,
                    &mut href_spans,
                    &mut section_source_depth,
                    &mut local_section_references,
                    xml.as_bytes()
                        .get(event_start..event_end)
                        .ok_or_else(|| invalid("ODM XML event span is outside content.xml"))?,
                    event_start,
                )?;
            },
            Event::Empty(element) => {
                let virtual_depth = checked_depth(depth)?;
                observe(
                    &reader,
                    namespace,
                    &element,
                    virtual_depth,
                    true,
                    &mut root_seen,
                    &mut body_seen,
                    &mut master_seen,
                    &mut body_depth,
                    &mut master_depth,
                    &mut section_names,
                    &mut xml_ids,
                    &mut sections,
                    &mut tree,
                    &mut references,
                    &mut href_spans,
                    &mut section_source_depth,
                    &mut local_section_references,
                    xml.as_bytes()
                        .get(event_start..event_end)
                        .ok_or_else(|| invalid("ODM XML event span is outside content.xml"))?,
                    event_start,
                )?;
            },
            Event::End(element) => {
                let local = element.local_name();
                if section_source_depth == Some(depth) {
                    section_source_depth = None;
                }
                if namespace == NamespaceKind::Text && local.as_ref() == b"section" {
                    let section = sections
                        .pop()
                        .ok_or_else(|| invalid("ODM text:section nesting underflow"))?;
                    if section.depth != depth {
                        return Err(invalid("ODM text:section nesting is malformed"));
                    }
                    tree.sections
                        .get_mut(section.position.get())
                        .ok_or_else(|| invalid("ODM section disappeared from its tree"))?
                        .source_span
                        .end = event_end;
                }
                if master_depth == Some(depth) {
                    master_depth = None;
                }
                if body_depth == Some(depth) {
                    body_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("ODM XML depth underflow"))?;
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODM content")),
            Event::GeneralRef(_) => {
                return Err(invalid("named XML entities are not allowed in ODM content"));
            },
            Event::Text(text) if section_source_depth.is_some() && !text.as_ref().is_empty() => {
                return Err(invalid("ODM text:section-source cannot contain text"));
            },
            Event::CData(text) if section_source_depth.is_some() && !text.as_ref().is_empty() => {
                return Err(invalid("ODM text:section-source cannot contain CDATA"));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_) => {},
        }
    }
    if depth != 0 || !root_seen || !body_seen || !master_seen {
        return Err(invalid(
            "ODM content has an incomplete master-document structure",
        ));
    }
    if !sections.is_empty() {
        return Err(invalid("ODM text:section nesting is incomplete"));
    }
    let mut positions = HashMap::new();
    positions
        .try_reserve(tree.sections.len())
        .map_err(|source| Error::Allocation {
            resource: "ODM local section target index",
            source,
        })?;
    for (index, section) in tree.sections.iter().enumerate() {
        positions.insert(section.name.as_str(), Position::new(index));
    }
    for reference in &mut tree.local_references {
        reference.target = positions.get(reference.target_name.as_str()).copied();
    }
    let structure = parse_master_structure(xml)?;
    Ok(Semantics {
        references,
        href_spans,
        tree,
        structure,
        local_section_references,
    })
}

fn parse_master_structure(xml: &str) -> Result<crate::structure::Structure> {
    const TABLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut depth = 0usize;
    let mut text_depth = None;
    let mut section_index = 0usize;
    let mut structure = crate::structure::Structure::default();
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid ODM master structure XML: {error}")))?;
        let is_office = matches!(&namespace, ResolveResult::Bound(uri) if uri.as_ref() == OFFICE);
        let is_text = matches!(&namespace, ResolveResult::Bound(uri) if uri.as_ref() == TEXT);
        let is_table = matches!(&namespace, ResolveResult::Bound(uri) if uri.as_ref() == TABLE);
        let event = event.into_owned();
        match event {
            Event::Start(element) => {
                depth = depth.saturating_add(1);
                observe_master_item(
                    element.local_name().as_ref(),
                    is_office,
                    is_text,
                    is_table,
                    depth,
                    &mut text_depth,
                    &mut section_index,
                    &mut structure,
                )?;
            },
            Event::Empty(element) => {
                observe_master_item(
                    element.local_name().as_ref(),
                    is_office,
                    is_text,
                    is_table,
                    depth.saturating_add(1),
                    &mut text_depth,
                    &mut section_index,
                    &mut structure,
                )?;
            },
            Event::End(element) => {
                if text_depth == Some(depth)
                    && is_office
                    && element.local_name().as_ref() == b"text"
                {
                    text_depth = None;
                }
                depth = depth.saturating_sub(1);
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in ODM content")),
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(structure)
}

#[allow(
    clippy::too_many_arguments,
    reason = "the XML event classifier keeps namespace and traversal state explicit"
)]
fn observe_master_item(
    local: &[u8],
    is_office: bool,
    is_text: bool,
    is_table: bool,
    depth: usize,
    text_depth: &mut Option<usize>,
    section_index: &mut usize,
    structure: &mut crate::structure::Structure,
) -> Result<()> {
    use crate::structure::{IndexKind, Kind};

    if is_office && local == b"text" {
        *text_depth = Some(depth);
    } else if *text_depth == Some(depth.saturating_sub(1)) {
        let kind = if is_text {
            match local {
                b"p" => Kind::Paragraph,
                b"h" => Kind::Heading,
                b"list" => Kind::List,
                b"section" => Kind::Section(Position::new(*section_index)),
                b"table-of-content" => Kind::GeneratedIndex(IndexKind::TableOfContents),
                b"illustration-index" => Kind::GeneratedIndex(IndexKind::Illustration),
                b"table-index" => Kind::GeneratedIndex(IndexKind::Table),
                b"object-index" => Kind::GeneratedIndex(IndexKind::Object),
                b"user-index" => Kind::GeneratedIndex(IndexKind::User),
                b"alphabetical-index" => Kind::GeneratedIndex(IndexKind::Alphabetical),
                b"bibliography" => Kind::GeneratedIndex(IndexKind::Bibliography),
                b"sequence-decls"
                | b"user-field-decls"
                | b"variable-decls"
                | b"dde-connection-decls"
                | b"tracked-changes" => Kind::Declarations,
                _ => Kind::Other,
            }
        } else if is_table && local == b"table" {
            Kind::Table
        } else {
            Kind::Other
        };
        if structure.items.len() >= MAX_IDENTITIES {
            return Err(invalid("ODM master-body item count exceeds the limit"));
        }
        structure
            .items
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODM master-body structure",
                source,
            })?;
        structure.items.push(kind);
    }
    if is_text && local == b"section" {
        *section_index = (*section_index).saturating_add(1);
    }
    Ok(())
}

#[allow(
    clippy::too_many_arguments,
    reason = "one XML event updates the bounded ODM semantic projection"
)]
fn observe(
    reader: &NsReader<&[u8]>,
    namespace: NamespaceKind,
    element: &BytesStart<'_>,
    depth: usize,
    empty: bool,
    root_seen: &mut bool,
    body_seen: &mut bool,
    master_seen: &mut bool,
    body_depth: &mut Option<usize>,
    master_depth: &mut Option<usize>,
    section_names: &mut HashSet<String>,
    xml_ids: &mut HashSet<String>,
    sections: &mut Vec<ActiveSection>,
    tree: &mut crate::model::section::Tree,
    references: &mut Vec<crate::model::subdocument::Reference>,
    href_spans: &mut Vec<Range<usize>>,
    section_source_depth: &mut Option<usize>,
    local_section_references: &mut Vec<(String, Range<usize>)>,
    tag: &[u8],
    tag_start: usize,
) -> Result<()> {
    let local = element.local_name();
    if section_source_depth.is_some() {
        return Err(invalid("ODM text:section-source cannot contain elements"));
    }
    let is_section_source = namespace == NamespaceKind::Text && local.as_ref() == b"section-source";
    let is_dde_source = namespace == NamespaceKind::Office && local.as_ref() == b"dde-source";
    if let Some(parent) = sections.last_mut()
        && depth == parent.depth.saturating_add(1)
    {
        if is_section_source || is_dde_source {
            if parent.child_seen || parent.source_seen {
                return Err(invalid(
                    "ODM linked-section source must occur once as the first section child",
                ));
            }
            parent.source_seen = true;
        } else {
            parent.child_seen = true;
        }
    }
    if depth == 1 {
        if *root_seen
            || namespace != NamespaceKind::Office
            || local.as_ref() != b"document-content"
            || empty
        {
            return Err(invalid(
                "ODM content requires one office:document-content root",
            ));
        }
        *root_seen = true;
    } else if namespace == NamespaceKind::Office && local.as_ref() == b"body" {
        if *body_seen || depth != 2 || empty {
            return Err(invalid("ODM content requires one non-empty office:body"));
        }
        *body_seen = true;
        *body_depth = Some(depth);
    } else if namespace == NamespaceKind::Office && local.as_ref() == b"text" {
        if *master_seen || *body_depth != Some(depth - 1) {
            return Err(invalid("ODM office:text body is misplaced or duplicated"));
        }
        *master_seen = true;
        if !empty {
            *master_depth = Some(depth);
        }
    } else if namespace == NamespaceKind::Text && local.as_ref() == b"section" {
        if master_depth.is_none() {
            return Err(invalid("ODM text:section is outside the master body"));
        }
        let name = attribute(reader, element, TEXT, b"name")?
            .ok_or_else(|| invalid("ODM text:section has no text:name"))?;
        ensure_short(&name, "ODM text:section name")?;
        let section = name.clone();
        insert_identity(section_names, name, "duplicate ODM section name")?;
        let style_name = attribute(reader, element, TEXT, b"style-name")?;
        let xml_id = attribute(reader, element, XML, b"id")?;
        let protected = attribute(reader, element, TEXT, b"protected")?
            .map(|value| parse_bool(&value, "ODM text:section text:protected"))
            .transpose()?;
        for (value, scope) in [
            (style_name.as_deref(), "ODM section style name"),
            (xml_id.as_deref(), "ODM section xml:id"),
        ] {
            if let Some(value) = value {
                ensure_short(value, scope)?;
            }
        }
        let position = Position::new(tree.sections.len());
        let name_key = attribute_key(reader, element, TEXT, b"name")?
            .ok_or_else(|| invalid("ODM section name source spelling disappeared"))?;
        let (name_start, name_end) = attribute_value_span(tag, &name_key)?;
        let parent = sections.last().map(|active| active.position);
        tree.sections
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODM section tree",
                source,
            })?;
        tree.sections.push(crate::model::section::Node {
            name: section.clone(),
            style_name,
            xml_id,
            protected,
            parent,
            children: Vec::new(),
            reference: None,
            local_reference: None,
            dde_source: false,
            source_span: tag_start..tag_start + tag.len(),
            name_span: tag_start + name_start..tag_start + name_end,
        });
        if let Some(parent_position) = parent {
            let parent_node = tree
                .sections
                .get_mut(parent_position.get())
                .ok_or_else(|| invalid("ODM section-tree parent disappeared"))?;
            parent_node
                .children
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODM section-tree children",
                    source,
                })?;
            parent_node.children.push(position);
        } else {
            tree.roots
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODM section-tree roots",
                    source,
                })?;
            tree.roots.push(position);
        }
        if !empty {
            sections
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODM section nesting",
                    source,
                })?;
            sections.push(ActiveSection {
                depth,
                position,
                name: section,
                child_seen: false,
                source_seen: false,
            });
        }
    } else if is_section_source {
        let Some(section) = sections.last() else {
            return Err(invalid("ODM text:section-source is outside a text:section"));
        };
        if depth != section.depth.saturating_add(1) {
            return Err(invalid(
                "ODM text:section-source is not a direct text:section child",
            ));
        }
        let href = attribute(reader, element, XLINK, b"href")?;
        let link_type = attribute(reader, element, XLINK, b"type")?;
        let show = attribute(reader, element, XLINK, b"show")?;
        if link_type.as_deref().is_some_and(|value| value != "simple") {
            return Err(invalid("ODM text:section-source xlink:type must be simple"));
        }
        if show.as_deref().is_some_and(|value| value != "embed") {
            return Err(invalid("ODM text:section-source xlink:show must be embed"));
        }
        if href.is_none() && (link_type.is_some() || show.is_some()) {
            return Err(invalid(
                "ODM text:section-source link attributes require xlink:href",
            ));
        }
        let source_section = attribute(reader, element, TEXT, b"section-name")?;
        let filter_name = attribute(reader, element, TEXT, b"filter-name")?;
        for (optional_value, scope) in [
            (source_section.as_deref(), "ODM source section name"),
            (filter_name.as_deref(), "ODM source filter name"),
        ] {
            if let Some(semantic_value) = optional_value {
                ensure_short(semantic_value, scope)?;
            }
        }
        if let Some(link_href) = href {
            ensure_short(&link_href, "ODM xlink:href")?;
            if references.len() >= MAX_REFERENCES {
                return Err(invalid("ODM subdocument reference count exceeds the limit"));
            }
            references
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODM subdocument references",
                    source,
                })?;
            href_spans
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODM subdocument reference spans",
                    source,
                })?;
            let href_key = attribute_key(reader, element, XLINK, b"href")?
                .ok_or_else(|| invalid("ODM xlink:href source spelling disappeared"))?;
            let (span_start, span_end) = attribute_value_span(tag, &href_key)?;
            references.push(crate::model::subdocument::Reference::new(
                section.name.clone(),
                link_href,
                source_section,
                filter_name,
            ));
            tree.sections
                .get_mut(section.position.get())
                .ok_or_else(|| invalid("ODM linked section disappeared from its tree"))?
                .reference = Some(Position::new(references.len() - 1));
            href_spans.push(tag_start + span_start..tag_start + span_end);
        } else if let Some(local_name) = source_section {
            let key = attribute_key(reader, element, TEXT, b"section-name")?
                .ok_or_else(|| invalid("ODM local section reference span is missing"))?;
            let (span_start, span_end) = attribute_value_span(tag, &key)?;
            local_section_references
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODM local section references",
                    source,
                })?;
            local_section_references
                .push((local_name, tag_start + span_start..tag_start + span_end));
            let reference_position = Position::new(tree.local_references.len());
            tree.local_references
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODM local section relationship projection",
                    source,
                })?;
            tree.local_references.push(crate::section::LocalReference {
                owner: section.position,
                target_name: local_section_references
                    .last()
                    .ok_or_else(|| invalid("ODM local section reference disappeared"))?
                    .0
                    .clone(),
                target: None,
            });
            tree.sections
                .get_mut(section.position.get())
                .ok_or_else(|| invalid("ODM local section owner disappeared"))?
                .local_reference = Some(reference_position);
        }
        if !empty {
            *section_source_depth = Some(depth);
        }
    } else if is_dde_source {
        let Some(section) = sections.last() else {
            return Err(invalid("ODM office:dde-source is outside a text:section"));
        };
        if depth != section.depth.saturating_add(1) {
            return Err(invalid(
                "ODM office:dde-source is not a direct text:section child",
            ));
        }
        tree.sections
            .get_mut(section.position.get())
            .ok_or_else(|| invalid("ODM DDE section owner disappeared"))?
            .dde_source = true;
    }
    if let Some(xml_id) = attribute(reader, element, XML, b"id")? {
        ensure_short(&xml_id, "ODM xml:id")?;
        insert_identity(xml_ids, xml_id, "duplicate ODM xml:id")?;
    }
    Ok(())
}

fn parse_bool(value: &str, scope: &str) -> Result<bool> {
    match value {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(invalid(format!("{scope} must be true or false"))),
    }
}

fn insert_identity(set: &mut HashSet<String>, value: String, message: &str) -> Result<()> {
    if set.len() >= MAX_IDENTITIES {
        return Err(invalid("ODM identity count exceeds the limit"));
    }
    set.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "ODM identities",
        source,
    })?;
    if !set.insert(value) {
        return Err(invalid(message));
    }
    Ok(())
}

fn ensure_short(value: &str, scope: &str) -> Result<()> {
    if value.len() > MAX_REFERENCE_BYTES {
        return Err(invalid(format!(
            "{scope} exceeds the {MAX_REFERENCE_BYTES}-byte limit"
        )));
    }
    Ok(())
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    let mut found_value = None;
    for raw_attribute in element.attributes() {
        let attribute =
            raw_attribute.map_err(|error| invalid(format!("invalid ODM attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
        if resolved_bound(&resolved, namespace) && name.as_ref() == local {
            let decoded = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(Cow::into_owned)
                .map_err(|error| invalid(format!("invalid ODM attribute value: {error}")))?;
            if found_value.replace(decoded).is_some() {
                return Err(invalid("duplicate namespace-equivalent ODM attribute"));
            }
        }
    }
    Ok(found_value)
}

fn attribute_key(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<Vec<u8>>> {
    let mut key = None;
    for raw_attribute in element.attributes() {
        let attribute =
            raw_attribute.map_err(|error| invalid(format!("invalid ODM attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
        if resolved_bound(&resolved, namespace)
            && name.as_ref() == local
            && key.replace(attribute.key.as_ref().to_vec()).is_some()
        {
            return Err(invalid("duplicate namespace-equivalent ODM attribute"));
        }
    }
    Ok(key)
}

fn attribute_value_span(tag: &[u8], wanted: &[u8]) -> Result<(usize, usize)> {
    let mut cursor = 1usize;
    while cursor < tag.len()
        && !tag[cursor].is_ascii_whitespace()
        && !matches!(tag[cursor], b'/' | b'>')
    {
        cursor += 1;
    }
    while cursor < tag.len() {
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || matches!(tag[cursor], b'/' | b'>') {
            break;
        }
        let name_start = cursor;
        while cursor < tag.len()
            && !tag[cursor].is_ascii_whitespace()
            && !matches!(tag[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return Err(invalid("ODM attribute is missing '='"));
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *tag
            .get(cursor)
            .filter(|quote| matches!(quote, b'\'' | b'"'))
            .ok_or_else(|| invalid("ODM attribute value is not quoted"))?;
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        if cursor >= tag.len() {
            return Err(invalid("ODM attribute value is unterminated"));
        }
        let value_end = cursor;
        cursor += 1;
        if &tag[name_start..name_end] == wanted {
            return Ok((value_start, value_end));
        }
    }
    Err(invalid("ODM attribute source span is missing"))
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_range_error| invalid("ODM XML source position exceeds the platform range"))
}

fn checked_depth(depth: usize) -> Result<usize> {
    let next_depth = depth
        .checked_add(1)
        .ok_or_else(|| invalid("ODM XML depth overflow"))?;
    if next_depth > MAX_DEPTH {
        return Err(invalid("ODM XML depth exceeds the limit"));
    }
    Ok(next_depth)
}

fn resolved_bound(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

fn classify(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == TEXT => NamespaceKind::Text,
        ResolveResult::Unbound | ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::validate;

    #[test]
    fn requires_family_body() {
        let text = concat!(
            r#"<office:document-content xmlns:office="#,
            r#""urn:oasis:names:tc:opendocument:xmlns:office:1.0">"#,
            "<office:body><office:text/></office:body>",
            "</office:document-content>",
        );
        let chart = concat!(
            r#"<office:document-content xmlns:office="#,
            r#""urn:oasis:names:tc:opendocument:xmlns:office:1.0">"#,
            "<office:body><office:chart/></office:body>",
            "</office:document-content>",
        );
        assert!(validate(text).is_ok());
        assert!(validate(chart).is_err());
    }
}
