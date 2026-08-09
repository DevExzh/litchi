//! Namespace-aware presentation-owner location and range rewriting.

use std::ops::Range;

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::{Limits, Tag, Tags};
use crate::shape::designer::{P202_NAMESPACE, TAGS_EXTENSION_URI};
use crate::{Error, Result};

const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const PML_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const P202: &[u8] = P202_NAMESPACE.as_bytes();

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Element {
    pub(crate) span: Range<usize>,
    pub(crate) start_end: usize,
    pub(crate) close_start: usize,
    pub(crate) empty: bool,
    pub(crate) qname: Vec<u8>,
    pub(crate) child_elements: usize,
    pub(crate) opaque_content: bool,
    pub(crate) non_namespace_attributes: usize,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Occurrence {
    pub(crate) outer: Element,
    pub(crate) payload: Element,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Layout {
    pub(crate) host: Element,
    pub(crate) extension_list: Option<Element>,
    pub(crate) occurrences: Vec<Occurrence>,
    pub(crate) relationship_id: String,
}

impl Layout {
    pub(crate) fn host_bytes<'a>(&self, source: &'a [u8]) -> &'a [u8] {
        source.get(self.host.span.clone()).unwrap_or_default()
    }
}

pub(crate) struct Located {
    pub(crate) tags: Vec<Tags>,
    pub(crate) layout: Layout,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Presentation,
    Designer,
    Other,
}

#[derive(Debug)]
struct Node {
    element: Element,
    namespace: NamespaceKind,
    local: Vec<u8>,
    parent: Option<usize>,
    id: Option<String>,
    relationship_id: Option<String>,
    uri: Option<String>,
    namespace_declarations: Vec<NamespaceDeclaration>,
    non_whitespace_text: bool,
}

#[derive(Debug)]
struct NamespaceDeclaration {
    prefix: Vec<u8>,
    uri: String,
}

pub(crate) fn locate(xml: &[u8], slide_id: u32, limits: Limits) -> Result<Located> {
    if xml.len() > limits.xml_bytes() {
        return Err(limit("Designer-tag owner XML bytes", limits.xml_bytes()));
    }
    if !(256..2_147_483_648).contains(&slide_id) {
        return Err(invalid("slide ID is outside ST_SlideId"));
    }
    let nodes = parse_nodes(xml, limits)?;
    let mut roots = nodes
        .iter()
        .enumerate()
        .filter(|(_, node)| node.parent.is_none());
    let (root, root_node) = roots
        .next()
        .ok_or_else(|| invalid("Designer-tag owner does not have one PresentationML root"))?;
    if roots.next().is_some()
        || root_node.namespace != NamespaceKind::Presentation
        || root_node.local.as_slice() != b"presentation"
    {
        return Err(invalid(
            "Designer-tag owner does not have one PresentationML root",
        ));
    }
    let mut lists = direct_children(&nodes, root)
        .filter(|(_, node)| is(node, NamespaceKind::Presentation, b"sldIdLst"));
    let list = lists
        .next()
        .map(|(index, _)| index)
        .ok_or_else(|| invalid("presentation does not contain the selected slide ID"))?;
    if lists.next().is_some() {
        return Err(invalid(
            "presentation has duplicate direct sldIdLst elements",
        ));
    }

    let mut selected = None;
    let mut seen_ids = std::collections::HashSet::new();
    seen_ids
        .try_reserve(direct_children(&nodes, list).count())
        .map_err(|source| Error::Allocation {
            resource: "Designer-tag slide-ID inventory",
            source,
        })?;
    for (index, node) in direct_children(&nodes, list)
        .filter(|(_, node)| is(node, NamespaceKind::Presentation, b"sldId"))
    {
        let raw_id = node
            .id
            .as_deref()
            .ok_or_else(|| invalid("p:sldId is missing its unqualified id attribute"))?;
        let id = raw_id
            .parse::<u32>()
            .map_err(|_err| invalid("p:sldId has an invalid numeric id"))?;
        if !(256..2_147_483_648).contains(&id) {
            return Err(invalid("p:sldId id is outside ST_SlideId"));
        }
        if !seen_ids.insert(id) {
            return Err(invalid("presentation contains duplicate slide IDs"));
        }
        if id == slide_id {
            selected = Some(index);
        }
    }
    let selected = selected.ok_or_else(|| invalid("selected slide ID was not found"))?;
    let host = &nodes[selected];
    let relationship_id = host
        .relationship_id
        .clone()
        .ok_or_else(|| invalid("selected p:sldId is missing its relationship ID"))?;
    let mut extension_lists = direct_children(&nodes, selected)
        .filter(|(_, node)| is(node, NamespaceKind::Presentation, b"extLst"));
    let extension_list = extension_lists.next();
    if extension_lists.next().is_some() {
        return Err(invalid(
            "selected p:sldId has duplicate direct extLst elements",
        ));
    }
    let mut occurrences = Vec::new();
    let mut tags = Vec::new();
    if let Some((list_index, list_node)) = extension_list {
        for (outer_index, outer) in direct_children(&nodes, list_index)
            .filter(|(_, node)| is(node, NamespaceKind::Presentation, b"ext"))
        {
            if outer.uri.as_deref().map(str::trim) != Some(TAGS_EXTENSION_URI) {
                continue;
            }
            let mut children = direct_children(&nodes, outer_index);
            let Some((payload_index, payload)) = children.next() else {
                return Err(invalid(
                    "Designer-tag extension does not contain exactly one p202:designTagLst",
                ));
            };
            if children.next().is_some()
                || !is(payload, NamespaceKind::Designer, b"designTagLst")
                || outer.non_whitespace_text
            {
                return Err(invalid(
                    "Designer-tag extension does not contain exactly one p202:designTagLst",
                ));
            }
            let bytes = xml
                .get(payload.element.span.clone())
                .ok_or_else(|| invalid("Designer-tag payload range is invalid"))?;
            let proof = designer_namespace_proof(&nodes, payload_index, limits)?;
            let value =
                crate::shape::designer::read_tags_with_prefix(bytes, limits, proof.as_deref())?;
            occurrences
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "Designer-tag extension inventory",
                    source,
                })?;
            tags.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "Designer-tag value inventory",
                source,
            })?;
            occurrences.push(Occurrence {
                outer: outer.element.clone(),
                payload: payload.element.clone(),
            });
            tags.push(value);
        }
        Ok(Located {
            tags,
            layout: Layout {
                host: host.element.clone(),
                extension_list: Some(list_node.element.clone()),
                occurrences,
                relationship_id,
            },
        })
    } else {
        Ok(Located {
            tags,
            layout: Layout {
                host: host.element.clone(),
                extension_list: None,
                occurrences,
                relationship_id,
            },
        })
    }
}

pub(crate) fn rewrite(
    source: &[u8],
    layout: &Layout,
    desired: Option<&Tags>,
    limits: Limits,
) -> Result<Vec<u8>> {
    if layout.occurrences.len() > 1 {
        return Err(super::model::ambiguous(layout.occurrences.len()));
    }
    match (layout.occurrences.first(), desired) {
        (Some(_), Some(value)) => {
            let payload = crate::shape::designer::write_tags(value, limits)?;
            replace(
                source,
                layout.occurrences[0].payload.span.clone(),
                &payload,
                limits,
            )
        },
        (Some(occurrence), None) => remove_occurrence(source, layout, occurrence, limits),
        (None, Some(value)) => insert_occurrence(source, layout, value, limits),
        (None, None) => copy_bounded(source, limits),
    }
}

pub(crate) fn clone_tags(value: &Tags, limits: Limits) -> Result<Tags> {
    value.validate(limits)?;
    let mut cloned = Tags::new();
    for tag in value.iter() {
        let name = clone_text(tag.name(), "Designer tag name clone")?;
        let value = clone_text(tag.value(), "Designer tag value clone")?;
        cloned.push_with_limits(Tag::new_with_limits(name, value, limits)?, limits)?;
    }
    Ok(cloned)
}

fn clone_text(value: &str, resource: &'static str) -> Result<String> {
    let mut cloned = String::new();
    cloned
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    cloned.push_str(value);
    Ok(cloned)
}

fn insert_occurrence(
    source: &[u8],
    layout: &Layout,
    value: &Tags,
    limits: Limits,
) -> Result<Vec<u8>> {
    let payload = crate::shape::designer::write_tags(value, limits)?;
    let ext_qname = qualified_sibling(
        layout
            .extension_list
            .as_ref()
            .map_or(layout.host.qname.as_slice(), |list| list.qname.as_slice()),
        b"ext",
    )?;
    let outer = build_outer(&ext_qname, &payload, limits)?;
    if let Some(list) = &layout.extension_list {
        if list.empty {
            return expand_empty(source, list, &outer, limits);
        }
        return replace(source, list.close_start..list.close_start, &outer, limits);
    }
    let list_qname = qualified_sibling(&layout.host.qname, b"extLst")?;
    let mut wrapper = Output::new(limits);
    wrapper.push(b"<")?;
    wrapper.push(&list_qname)?;
    wrapper.push(b">")?;
    wrapper.push(&outer)?;
    wrapper.push(b"</")?;
    wrapper.push(&list_qname)?;
    wrapper.push(b">")?;
    let wrapper = wrapper.finish();
    if layout.host.empty {
        expand_empty(source, &layout.host, &wrapper, limits)
    } else {
        replace(
            source,
            layout.host.close_start..layout.host.close_start,
            &wrapper,
            limits,
        )
    }
}

fn remove_occurrence(
    source: &[u8],
    layout: &Layout,
    occurrence: &Occurrence,
    limits: Limits,
) -> Result<Vec<u8>> {
    let list = layout
        .extension_list
        .as_ref()
        .ok_or_else(|| invalid("Designer-tag extension has no owning extLst"))?;
    if list.child_elements == 1
        && !list.opaque_content
        && list.non_namespace_attributes == 0
        && layout.host.child_elements == 1
        && !layout.host.opaque_content
    {
        let without_list = replace(source, list.span.clone(), &[], limits)?;
        let removed = list.span.end - list.span.start;
        let adjusted = Element {
            span: layout.host.span.start..layout.host.span.end - removed,
            start_end: layout.host.start_end,
            close_start: layout.host.close_start - removed,
            empty: false,
            qname: layout.host.qname.clone(),
            child_elements: 0,
            opaque_content: false,
            non_namespace_attributes: layout.host.non_namespace_attributes,
        };
        return collapse_empty(&without_list, &adjusted, limits);
    }
    replace(source, occurrence.outer.span.clone(), &[], limits)
}

fn parse_nodes(xml: &[u8], limits: Limits) -> Result<Vec<Node>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut nodes = Vec::<Node>::new();
    let mut stack = Vec::<usize>::new();
    let mut root_closed = false;
    let mut events = 0usize;
    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let namespace = namespace_kind(namespace);
        let event = event.into_owned();
        if !matches!(event, Event::Eof) {
            events = events
                .checked_add(1)
                .ok_or_else(|| limit("Designer-tag owner XML events", limits.xml_nodes()))?;
            if events > limits.xml_nodes() {
                return Err(limit("Designer-tag owner XML events", limits.xml_nodes()));
            }
        }
        let empty = matches!(&event, Event::Empty(_));
        let end = position(&reader)?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if root_closed && stack.is_empty() {
                    return Err(invalid("presentation XML contains trailing elements"));
                }
                if stack.len() >= limits.xml_depth() {
                    return Err(limit("Designer-tag owner XML depth", limits.xml_depth()));
                }
                nodes.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "Designer-tag owner XML nodes",
                    source,
                })?;
                let parent = stack.last().copied();
                let qname = copy_name(element.name().as_ref(), limits)?;
                let local = copy_name(element.local_name().as_ref(), limits)?;
                let attributes = attributes(&element, decoder, &reader, limits)?;
                let index = nodes.len();
                nodes.push(Node {
                    element: Element {
                        span: start..end,
                        start_end: end,
                        close_start: end,
                        empty,
                        qname,
                        child_elements: 0,
                        opaque_content: false,
                        non_namespace_attributes: attributes.non_namespace,
                    },
                    namespace,
                    local,
                    parent,
                    id: attributes.id,
                    relationship_id: attributes.relationship_id,
                    uri: attributes.uri,
                    namespace_declarations: attributes.namespace_declarations,
                    non_whitespace_text: false,
                });
                if let Some(parent) = parent {
                    nodes[parent].element.child_elements = nodes[parent]
                        .element
                        .child_elements
                        .checked_add(1)
                        .ok_or_else(|| limit("Designer-tag owner XML nodes", limits.xml_nodes()))?;
                }
                if !empty {
                    stack.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "Designer-tag owner XML stack",
                        source,
                    })?;
                    stack.push(index);
                } else if parent.is_none() {
                    root_closed = true;
                }
            },
            Event::End(_) => {
                let index = stack
                    .pop()
                    .ok_or_else(|| invalid("presentation XML stack underflow"))?;
                nodes[index].element.close_start = start;
                nodes[index].element.span.end = end;
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Text(text) => {
                let non_whitespace = !text
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .trim()
                    .is_empty();
                if let Some(index) = stack.last().copied() {
                    nodes[index].non_whitespace_text |= non_whitespace;
                    nodes[index].element.opaque_content |= non_whitespace;
                } else if non_whitespace {
                    return Err(invalid("presentation XML has text outside its root"));
                }
            },
            Event::Comment(_) => {
                if let Some(index) = stack.last().copied() {
                    nodes[index].element.opaque_content = true;
                }
            },
            Event::CData(value) => {
                if let Some(index) = stack.last().copied() {
                    let non_whitespace = value
                        .as_ref()
                        .iter()
                        .any(|byte| !byte.is_ascii_whitespace());
                    nodes[index].non_whitespace_text |= non_whitespace;
                    nodes[index].element.opaque_content = true;
                }
            },
            Event::Decl(_) if nodes.is_empty() && stack.is_empty() && !root_closed => {},
            Event::Decl(_) => {
                return Err(invalid(
                    "Designer-tag owner XML has a misplaced XML declaration",
                ));
            },
            Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_) => {
                return Err(invalid(
                    "Designer-tag owner XML contains forbidden DTD, PI, or entity markup",
                ));
            },
            Event::Eof => break,
        }
    }
    if !stack.is_empty() || !root_closed {
        return Err(invalid("presentation XML is unterminated"));
    }
    Ok(nodes)
}

struct Attributes {
    id: Option<String>,
    relationship_id: Option<String>,
    uri: Option<String>,
    namespace_declarations: Vec<NamespaceDeclaration>,
    non_namespace: usize,
}

fn attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    reader: &NsReader<&[u8]>,
    limits: Limits,
) -> Result<Attributes> {
    let mut id = None;
    let mut uri = None;
    let mut namespace_declarations = Vec::new();
    let mut non_namespace = 0usize;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_ref().len() > limits.attribute_bytes()
            || attribute.value.len() > limits.attribute_bytes()
        {
            return Err(limit(
                "Designer-tag owner XML attribute bytes",
                limits.attribute_bytes(),
            ));
        }
        let key = attribute.key.as_ref();
        let namespace = key == b"xmlns" || key.starts_with(b"xmlns:");
        if !namespace {
            non_namespace = non_namespace
                .checked_add(1)
                .ok_or_else(|| limit("Designer-tag owner XML attributes", limits.xml_nodes()))?;
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        if value.len() > limits.attribute_bytes() {
            return Err(limit(
                "Designer-tag owner XML attribute bytes",
                limits.attribute_bytes(),
            ));
        }
        if namespace {
            let prefix = if key == b"xmlns" {
                Vec::new()
            } else {
                copy_name(&key[b"xmlns:".len()..], limits)?
            };
            namespace_declarations
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "Designer-tag namespace declarations",
                    source,
                })?;
            namespace_declarations.push(NamespaceDeclaration {
                prefix,
                uri: value.into_owned(),
            });
        } else if key == b"id" {
            if id.replace(value.into_owned()).is_some() {
                return Err(invalid("element has duplicate unqualified id attributes"));
            }
        } else if key == b"uri" && uri.replace(value.into_owned()).is_some() {
            return Err(invalid("element has duplicate unqualified uri attributes"));
        }
    }
    let relationship_id =
        crate::namespace::relationship_attribute_value(element, b"id", decoder, reader.resolver())?;
    Ok(Attributes {
        id,
        relationship_id,
        uri,
        namespace_declarations,
        non_namespace,
    })
}

fn direct_children(nodes: &[Node], parent: usize) -> impl Iterator<Item = (usize, &Node)> {
    nodes
        .iter()
        .enumerate()
        .filter(move |(_, node)| node.parent == Some(parent))
}

fn is(node: &Node, namespace: NamespaceKind, local: &[u8]) -> bool {
    node.namespace == namespace && node.local.as_slice() == local
}

fn namespace_kind(namespace: ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if value == PML || value == PML_STRICT => {
            NamespaceKind::Presentation
        },
        ResolveResult::Bound(Namespace(value)) if value == P202 => NamespaceKind::Designer,
        _ => NamespaceKind::Other,
    }
}

fn designer_namespace_proof(
    nodes: &[Node],
    node: usize,
    limits: Limits,
) -> Result<Option<Vec<u8>>> {
    let mut lineage = Vec::new();
    let mut current = Some(node);
    while let Some(index) = current {
        lineage.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "Designer-tag namespace lineage",
            source,
        })?;
        lineage.push(index);
        current = nodes[index].parent;
    }

    let mut bindings: Vec<(&[u8], &str)> = Vec::new();
    for index in lineage.into_iter().rev() {
        for declaration in &nodes[index].namespace_declarations {
            if let Some(binding) = bindings
                .iter_mut()
                .find(|(prefix, _)| *prefix == declaration.prefix.as_slice())
            {
                binding.1 = declaration.uri.as_str();
            } else {
                bindings
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "Designer-tag namespace bindings",
                        source,
                    })?;
                bindings.push((declaration.prefix.as_slice(), declaration.uri.as_str()));
            }
        }
    }

    let matching = bindings
        .iter()
        .filter(|(_, uri)| uri.as_bytes() == P202)
        .count();
    if matching == 0 {
        return Ok(None);
    }
    let mut proof = Vec::new();
    for (prefix, _) in bindings
        .into_iter()
        .filter(|(_, uri)| uri.as_bytes() == P202)
    {
        if !proof.is_empty() {
            proof.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "Designer-tag namespace proof",
                source,
            })?;
            proof.push(0);
        } else if matching > 1 && prefix.is_empty() {
            proof.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "Designer-tag namespace proof",
                source,
            })?;
            proof.push(0);
        }
        proof
            .try_reserve(prefix.len())
            .map_err(|source| Error::Allocation {
                resource: "Designer-tag namespace proof",
                source,
            })?;
        proof.extend_from_slice(prefix);
        if proof.len() > limits.xml_bytes() {
            return Err(limit(
                "Designer-tag namespace proof bytes",
                limits.xml_bytes(),
            ));
        }
    }
    Ok(Some(proof))
}

fn build_outer(qname: &[u8], payload: &[u8], limits: Limits) -> Result<Vec<u8>> {
    let mut output = Output::new(limits);
    output.push(b"<")?;
    output.push(qname)?;
    output.push(b" uri=\"")?;
    output.push(TAGS_EXTENSION_URI.as_bytes())?;
    output.push(b"\">")?;
    output.push(payload)?;
    output.push(b"</")?;
    output.push(qname)?;
    output.push(b">")?;
    Ok(output.finish())
}

fn qualified_sibling(qname: &[u8], local: &[u8]) -> Result<Vec<u8>> {
    let prefix = qname.iter().position(|byte| *byte == b':');
    let length = prefix.map_or(local.len(), |position| position + 1 + local.len());
    let mut result = Vec::new();
    result
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "Designer-tag qualified name",
            source,
        })?;
    if let Some(position) = prefix {
        result.extend_from_slice(&qname[..=position]);
    }
    result.extend_from_slice(local);
    Ok(result)
}

fn expand_empty(source: &[u8], element: &Element, child: &[u8], limits: Limits) -> Result<Vec<u8>> {
    let raw = source
        .get(element.span.clone())
        .ok_or_else(|| invalid("Designer-tag empty-element range is invalid"))?;
    let slash = raw
        .iter()
        .rposition(|byte| *byte == b'/')
        .ok_or_else(|| invalid("Designer-tag empty element has no closing slash"))?;
    let mut output = Output::new(limits);
    output.push(&raw[..slash])?;
    output.push(&raw[slash + 1..])?;
    output.push(child)?;
    output.push(b"</")?;
    output.push(&element.qname)?;
    output.push(b">")?;
    replace(source, element.span.clone(), &output.finish(), limits)
}

fn collapse_empty(source: &[u8], element: &Element, limits: Limits) -> Result<Vec<u8>> {
    let raw = source
        .get(element.span.clone())
        .ok_or_else(|| invalid("Designer-tag collapse range is invalid"))?;
    let start_tag = raw
        .get(..element.start_end - element.span.start)
        .ok_or_else(|| invalid("Designer-tag start-tag range is invalid"))?;
    let mut replacement = Vec::new();
    replacement
        .try_reserve_exact(
            start_tag
                .len()
                .checked_add(1)
                .ok_or_else(|| limit("Designer-tag output bytes", limits.xml_bytes()))?,
        )
        .map_err(|source| Error::Allocation {
            resource: "Designer-tag empty-element collapse",
            source,
        })?;
    replacement.extend_from_slice(&start_tag[..start_tag.len() - 1]);
    replacement.extend_from_slice(b"/>");
    replace(source, element.span.clone(), &replacement, limits)
}

fn replace(source: &[u8], range: Range<usize>, value: &[u8], limits: Limits) -> Result<Vec<u8>> {
    let result = replace_unbounded(source, range, value)?;
    if result.len() > limits.xml_bytes() {
        return Err(limit("Designer-tag output XML bytes", limits.xml_bytes()));
    }
    Ok(result)
}

fn copy_bounded(source: &[u8], limits: Limits) -> Result<Vec<u8>> {
    replace(source, 0..0, &[], limits)
}

fn replace_unbounded(source: &[u8], range: Range<usize>, value: &[u8]) -> Result<Vec<u8>> {
    if range.start > range.end || range.end > source.len() {
        return Err(invalid("Designer-tag replacement range is invalid"));
    }
    let length = source
        .len()
        .checked_sub(range.end - range.start)
        .and_then(|length| length.checked_add(value.len()))
        .ok_or_else(|| limit("Designer-tag output XML bytes", usize::MAX))?;
    let mut result = Vec::new();
    result
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "Designer-tag output XML",
            source,
        })?;
    result.extend_from_slice(&source[..range.start]);
    result.extend_from_slice(value);
    result.extend_from_slice(&source[range.end..]);
    Ok(result)
}

fn copy_name(value: &[u8], limits: Limits) -> Result<Vec<u8>> {
    if value.is_empty() || value.len() > limits.attribute_bytes() {
        return Err(limit(
            "Designer-tag owner XML name bytes",
            limits.attribute_bytes(),
        ));
    }
    let mut result = Vec::new();
    result
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation {
            resource: "Designer-tag owner XML name",
            source,
        })?;
    result.extend_from_slice(value);
    Ok(result)
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_err| invalid("Designer-tag owner XML position does not fit usize"))
}

struct Output {
    bytes: Vec<u8>,
    limits: Limits,
}

impl Output {
    fn new(limits: Limits) -> Self {
        Self {
            bytes: Vec::new(),
            limits,
        }
    }

    fn push(&mut self, value: &[u8]) -> Result<()> {
        let length = self
            .bytes
            .len()
            .checked_add(value.len())
            .ok_or_else(|| limit("Designer-tag output XML bytes", self.limits.xml_bytes()))?;
        if length > self.limits.xml_bytes() {
            return Err(limit(
                "Designer-tag output XML bytes",
                self.limits.xml_bytes(),
            ));
        }
        self.bytes
            .try_reserve(value.len())
            .map_err(|source| Error::Allocation {
                resource: "Designer-tag output XML",
                source,
            })?;
        self.bytes.extend_from_slice(value);
        Ok(())
    }

    fn finish(self) -> Vec<u8> {
        self.bytes
    }
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
