use std::ops::Range;

use crate::presentation::embedded::{
    MAX_XML_ATTRIBUTES, MAX_XML_BYTES, MAX_XML_DEPTH, PML, REL, STRICT_PML, STRICT_REL, bounded,
    increment_nodes, invalid, limit,
};
use crate::{Error, Result};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::super::codec::OLE_GRAPHIC_DATA_URI;

const MC: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

#[derive(Debug, Clone)]
pub(crate) struct Attribute {
    pub(crate) local: Vec<u8>,
    pub(crate) namespace: Vec<u8>,
    pub(crate) name_start: usize,
    pub(crate) value_start: usize,
    pub(crate) value_end: usize,
    pub(crate) value: Vec<u8>,
}

#[derive(Debug, Clone)]
pub(crate) struct Node {
    pub(crate) local: Vec<u8>,
    pub(crate) namespace: Vec<u8>,
    pub(crate) attributes: Vec<Attribute>,
    pub(crate) children: Vec<Node>,
    pub(crate) start: usize,
    pub(crate) open_end: usize,
    pub(crate) close_start: Option<usize>,
    pub(crate) end: usize,
}

#[derive(Debug, Clone)]
pub(crate) struct Located {
    pub(crate) frame: Range<usize>,
    pub(crate) object: Node,
    pub(crate) anchor: Option<(Node, Node)>,
}

#[derive(Debug, Clone)]
pub(crate) struct Document {
    pub(crate) frames: Vec<Located>,
    pub(crate) insertion: usize,
    pub(crate) max_shape_id: u32,
}

#[derive(Debug, Clone)]
pub(crate) struct Edit {
    pub(crate) range: Range<usize>,
    pub(crate) replacement: Vec<u8>,
}

pub(crate) fn locate(source: &[u8]) -> Result<Document> {
    if source.len() > MAX_XML_BYTES {
        return Err(limit("OLE slide XML bytes", MAX_XML_BYTES));
    }
    let root = parse(source)?;
    let insertion = find_shape_tree(&root)
        .and_then(|node| node.close_start)
        .ok_or_else(|| invalid("OLE slide has no non-empty shape tree"))?;
    let mut frames = Vec::new();
    collect_frames(&root, &mut frames)?;
    if frames.len() > super::validation::MAX_OBJECTS {
        return Err(limit("OLE object count", super::validation::MAX_OBJECTS));
    }
    Ok(Document {
        frames,
        insertion,
        max_shape_id: max_shape_id(&root)?,
    })
}

fn parse(source: &[u8]) -> Result<Node> {
    let mut reader = NsReader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut stack = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut root_seen = false;
    loop {
        let before = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let after = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                if stack.len() >= MAX_XML_DEPTH {
                    return Err(limit("OLE XML depth", MAX_XML_DEPTH));
                }
                let node = make_node(
                    source, &element, before, after, decoder, &resolver, &namespace,
                )?;
                if stack.is_empty() {
                    if root_seen {
                        return Err(invalid("OLE slide has multiple roots"));
                    }
                    root_seen = true;
                }
                stack.push(node);
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                let node = make_node(
                    source, &element, before, after, decoder, &resolver, &namespace,
                )?;
                attach(node, &mut stack, &mut root)?;
            },
            Event::End(element) => {
                let mut node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected OLE closing element"))?;
                if node.local.as_slice() != element.local_name().as_ref() {
                    return Err(invalid("mismatched OLE closing element"));
                }
                node.close_start = Some(before);
                node.end = after;
                attach(node, &mut stack, &mut root)?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "OLE slide rejects DTDs and processing instructions",
                ));
            },
            Event::Text(_) | Event::CData(_) | Event::Comment(_) => {},
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated OLE slide XML"));
    }
    let root = root.ok_or_else(|| invalid("OLE slide has no root"))?;
    if root.local.as_slice() != b"sld"
        || (root.namespace.as_slice() != PML && root.namespace.as_slice() != STRICT_PML)
    {
        return Err(invalid("OLE slide XML must have a PresentationML sld root"));
    }
    Ok(root)
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_err| invalid("OLE slide XML offset overflow"))
}

fn make_node(
    source: &[u8],
    element: &BytesStart<'_>,
    start: usize,
    open_end: usize,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    namespace: &ResolveResult<'_>,
) -> Result<Node> {
    if element.attributes().with_checks(true).count() > MAX_XML_ATTRIBUTES {
        return Err(limit("OLE XML attributes", MAX_XML_ATTRIBUTES));
    }
    let namespace = resolve_namespace(namespace);
    let attributes = scan_attributes(source, start, open_end, decoder, resolver)?;
    Ok(Node {
        local: element.local_name().as_ref().to_vec(),
        namespace,
        attributes,
        children: Vec::new(),
        start,
        open_end,
        close_start: None,
        end: open_end,
    })
}

fn scan_attributes(
    source: &[u8],
    start: usize,
    open_end: usize,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Vec<Attribute>> {
    let end = open_end
        .checked_sub(1)
        .ok_or_else(|| invalid("OLE XML opening tag is truncated"))?;
    let mut cursor = start
        .checked_add(1)
        .ok_or_else(|| invalid("OLE XML offset overflow"))?;
    while cursor < end && !is_space(source[cursor]) && source[cursor] != b'/' {
        cursor += 1;
    }
    let mut attributes = Vec::new();
    while cursor < end {
        while cursor < end && is_space(source[cursor]) {
            cursor += 1;
        }
        if cursor >= end || source[cursor] == b'/' {
            break;
        }
        let name_start = cursor;
        while cursor < end
            && !is_space(source[cursor])
            && !matches!(source[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        if name_start == cursor {
            return Err(invalid("OLE XML attribute has no name"));
        }
        let qualified = source[name_start..cursor].to_vec();
        while cursor < end && is_space(source[cursor]) {
            cursor += 1;
        }
        if cursor >= end || source[cursor] != b'=' {
            return Err(invalid("OLE XML attribute has no value"));
        }
        cursor += 1;
        while cursor < end && is_space(source[cursor]) {
            cursor += 1;
        }
        if cursor >= end || !matches!(source[cursor], b'\'' | b'"') {
            return Err(invalid("OLE XML attribute value is not quoted"));
        }
        let quote = source[cursor];
        cursor += 1;
        let value_start = cursor;
        while cursor < end && source[cursor] != quote {
            cursor += 1;
        }
        if cursor >= end {
            return Err(invalid("OLE XML attribute value is unterminated"));
        }
        let value_end = cursor;
        cursor += 1;
        let key = quick_xml::name::QName(&qualified);
        let (namespace, _) = resolver.resolve_attribute(key);
        let namespace = resolve_namespace(&namespace);
        let local = element_local(&qualified);
        let value = decoder
            .decode(&source[value_start..value_end])
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        bounded(&value, "OLE XML attribute")?;
        attributes.push(Attribute {
            local,
            namespace,
            name_start,
            value_start,
            value_end,
            value: value.into_bytes(),
        });
    }
    Ok(attributes)
}

fn resolve_namespace(namespace: &ResolveResult<'_>) -> Vec<u8> {
    match namespace {
        ResolveResult::Bound(value) => value.0.to_vec(),
        ResolveResult::Unknown(prefix) => prefix.as_slice().to_vec(),
        ResolveResult::Unbound => Vec::new(),
    }
}

fn element_local(name: &[u8]) -> Vec<u8> {
    name.rsplit(|byte| *byte == b':')
        .next()
        .unwrap_or(name)
        .to_vec()
}

fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("OLE slide has multiple roots"));
    }
    Ok(())
}

fn collect_frames(node: &Node, frames: &mut Vec<Located>) -> Result<()> {
    if node.local.as_slice() == b"graphicFrame"
        && let Some((object, anchor)) = find_object(node)?
    {
        frames.push(Located {
            frame: node.start..node.end,
            object,
            anchor,
        });
    }
    for child in &node.children {
        collect_frames(child, frames)?;
    }
    Ok(())
}

fn find_object(frame: &Node) -> Result<Option<(Node, Option<(Node, Node)>)>> {
    let Some(graphic) = child(frame, b"graphic") else {
        return Ok(None);
    };
    let Some(data) = child(graphic, b"graphicData") else {
        return Ok(None);
    };
    if attribute(data, b"uri", false)
        .is_none_or(|value| value.as_slice() != OLE_GRAPHIC_DATA_URI.as_bytes())
    {
        return Ok(None);
    }
    let Some(object) = descendant(data, b"oleObj") else {
        return Ok(None);
    };
    let anchor = child(frame, b"xfrm")
        .and_then(|xfrm| Some((child(xfrm, b"off")?.clone(), child(xfrm, b"ext")?.clone())));
    Ok(Some((object.clone(), anchor)))
}

fn child<'a>(node: &'a Node, local: &[u8]) -> Option<&'a Node> {
    active_children(node)
        .into_iter()
        .find(|child| child.local.as_slice() == local)
}

fn descendant<'a>(node: &'a Node, local: &[u8]) -> Option<&'a Node> {
    if node.local.as_slice() == local {
        return Some(node);
    }
    for child in active_children(node) {
        if let Some(found) = descendant(child, local) {
            return Some(found);
        }
    }
    None
}

fn active_children(node: &Node) -> Vec<&Node> {
    let mut result = Vec::new();
    for child in &node.children {
        if child.local.as_slice() == b"AlternateContent" && child.namespace.as_slice() == MC {
            if let Some(fallback) = child
                .children
                .iter()
                .find(|value| value.local.as_slice() == b"Fallback")
            {
                result.extend(fallback.children.iter());
            } else if let Some(choice) = child
                .children
                .iter()
                .find(|value| value.local.as_slice() == b"Choice")
            {
                result.extend(choice.children.iter());
            }
        } else {
            result.push(child);
        }
    }
    result
}

fn attribute(node: &Node, local: &[u8], relationship: bool) -> Option<Vec<u8>> {
    node.attributes.iter().find_map(|attribute| {
        if attribute.local.as_slice() != local {
            return None;
        }
        let allowed = if relationship {
            attribute.namespace.as_slice() == REL
                || attribute.namespace.as_slice() == STRICT_REL
                || attribute.namespace.as_slice() == b"r"
        } else {
            attribute.namespace.is_empty()
        };
        allowed.then(|| attribute_value(node, attribute))
    })
}

fn attribute_value(node: &Node, attribute: &Attribute) -> Vec<u8> {
    let _ = node;
    attribute.value.clone()
}

fn find_shape_tree(node: &Node) -> Option<&Node> {
    if node.local.as_slice() == b"spTree" {
        return Some(node);
    }
    for child in &node.children {
        if let Some(found) = find_shape_tree(child) {
            return Some(found);
        }
    }
    None
}

fn max_shape_id(root: &Node) -> Result<u32> {
    let mut maximum = 1u32;
    visit(root, &mut |node| {
        if node.local.as_slice() == b"cNvPr"
            && let Some(value) = attribute(node, b"id", false)
        {
            let value = String::from_utf8_lossy(&value);
            let parsed = value
                .parse::<u32>()
                .map_err(|_err| invalid("invalid PresentationML shape ID"))?;
            maximum = maximum.max(parsed);
        }
        Ok(())
    })?;
    Ok(maximum)
}

fn visit(node: &Node, visitor: &mut impl FnMut(&Node) -> Result<()>) -> Result<()> {
    visitor(node)?;
    for child in &node.children {
        visit(child, visitor)?;
    }
    Ok(())
}

pub(crate) fn attribute_edit(
    source: &[u8],
    node: &Node,
    local: &[u8],
    relationship: bool,
    value: Option<&str>,
) -> Result<Option<Edit>> {
    let found = node.attributes.iter().find(|attribute| {
        if attribute.local.as_slice() != local {
            return false;
        }
        if relationship {
            attribute.namespace.as_slice() == REL
                || attribute.namespace.as_slice() == STRICT_REL
                || attribute.namespace.as_slice() == b"r"
        } else {
            attribute.namespace.is_empty()
        }
    });
    match (found, value) {
        (Some(attribute), Some(value)) => Ok(Some(Edit {
            range: attribute.value_start..attribute.value_end,
            replacement: escape(value).into_bytes(),
        })),
        (Some(attribute), None) => {
            let mut start = attribute.name_start;
            while start > node.start + 1 && is_space(source[start - 1]) {
                start -= 1;
            }
            Ok(Some(Edit {
                range: start..attribute.value_end + 1,
                replacement: Vec::new(),
            }))
        },
        (None, Some(value)) => {
            let insert = node
                .open_end
                .checked_sub(1)
                .ok_or_else(|| invalid("OLE XML opening tag is truncated"))?;
            let insert = if source.get(insert) == Some(&b'/') {
                insert
            } else {
                insert
            };
            let name = if relationship {
                format!("r:{}", String::from_utf8_lossy(local))
            } else {
                String::from_utf8_lossy(local).into_owned()
            };
            Ok(Some(Edit {
                range: insert..insert,
                replacement: format!(" {name}=\"{}\"", escape(value)).into_bytes(),
            }))
        },
        (None, None) => Ok(None),
    }
}

pub(crate) fn apply_edits(source: &[u8], mut edits: Vec<Edit>) -> Result<Vec<u8>> {
    if edits.is_empty() {
        return Ok(source.to_vec());
    }
    edits.sort_by(|left, right| {
        right
            .range
            .start
            .cmp(&left.range.start)
            .then_with(|| right.range.end.cmp(&left.range.end))
    });
    for pair in edits.windows(2) {
        if pair[0].range.start < pair[1].range.end {
            return Err(invalid("OLE source edits overlap"));
        }
    }
    let mut output = source.to_vec();
    for edit in edits {
        if edit.range.end > output.len() {
            return Err(invalid("OLE source edit escapes the slide"));
        }
        output.splice(edit.range, edit.replacement);
    }
    if output.len() > MAX_XML_BYTES {
        return Err(limit("updated OLE slide XML bytes", MAX_XML_BYTES));
    }
    Ok(output)
}

pub(crate) fn append_fragment(source: &[u8], insertion: usize, fragment: &[u8]) -> Result<Edit> {
    if insertion > source.len() {
        return Err(invalid("OLE insertion point escapes the slide"));
    }
    Ok(Edit {
        range: insertion..insertion,
        replacement: fragment.to_vec(),
    })
}

fn is_space(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\n' | b'\r')
}

fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}
