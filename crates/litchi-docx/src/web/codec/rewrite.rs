//! Source-preserving rewrites for the small web-settings edit vocabulary.
//!
//! The semantic parser remains the authority for the complete part.  This
//! module only locates the selected known element and replaces that element
//! (or inserts/removes it); every other byte, including extension elements,
//! comments, namespace declarations, and producer whitespace, is copied from
//! the source.

use std::fmt::Write as _;

use quick_xml::XmlVersion;
use quick_xml::escape::escape;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

use super::super::model::{Border, BorderSide, Id, Layout, Screen};
use super::super::transaction::Edit;
use super::super::{MAX_XML_BYTES, MAX_XML_EVENTS, Result, invalid, is_wordprocessing_namespace};
use crate::Error;

#[derive(Debug, Clone)]
struct Node {
    start: usize,
    open_end: usize,
    close_start: usize,
    end: usize,
    children: Vec<usize>,
    local: Vec<u8>,
    qname: Vec<u8>,
    word: bool,
    empty: bool,
}

#[derive(Debug)]
struct Tree {
    nodes: Vec<Node>,
    root: usize,
}

impl Tree {
    fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_XML_BYTES {
            return Err(invalid(format!(
                "web-settings XML exceeds {MAX_XML_BYTES} bytes"
            )));
        }

        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().check_end_names = true;
        let mut nodes = Vec::new();
        let mut stack = Vec::new();
        let mut root = None;

        loop {
            let event_start = reader.buffer_position() as usize;
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            let word = is_wordprocessing_namespace(&namespace);
            let event = event.into_owned();
            let event_end = reader.buffer_position() as usize;
            match event {
                Event::Start(element) => {
                    let parent = stack.last().copied();
                    let index = push_node(
                        &mut nodes,
                        event_start,
                        event_end,
                        parent,
                        element,
                        word,
                        false,
                    )?;
                    if root.is_none() {
                        root = Some(index);
                    }
                    stack.push(index);
                },
                Event::Empty(element) => {
                    let parent = stack.last().copied();
                    let index = push_node(
                        &mut nodes,
                        event_start,
                        event_end,
                        parent,
                        element,
                        word,
                        true,
                    )?;
                    if root.is_none() {
                        root = Some(index);
                    }
                },
                Event::End(_) => {
                    let index = stack
                        .pop()
                        .ok_or_else(|| invalid("web-settings XML has an unexpected end tag"))?;
                    nodes[index].close_start = event_start;
                    nodes[index].end = event_end;
                },
                Event::Eof => break,
                _ => {},
            }
        }

        if !stack.is_empty() {
            return Err(invalid("web-settings XML has an unterminated element"));
        }
        let root = root.ok_or_else(|| invalid("web-settings XML has no root element"))?;
        if nodes[root].local.as_slice() != b"webSettings" || !nodes[root].word {
            return Err(invalid("web-settings XML root is not w:webSettings"));
        }
        Ok(Self { nodes, root })
    }

    fn child(&self, parent: usize, name: &[u8]) -> Result<Option<usize>> {
        let mut found = None;
        for &index in &self.nodes[parent].children {
            let node = &self.nodes[index];
            if node.word && node.local.as_slice() == name {
                if found.is_some() {
                    return Err(invalid(format!(
                        "web-settings element '{}' occurs more than once",
                        String::from_utf8_lossy(name)
                    )));
                }
                found = Some(index);
            }
        }
        Ok(found)
    }

    fn child_at_rank(&self, parent: usize, rank: u8) -> Option<usize> {
        self.nodes[parent].children.iter().copied().find(|&index| {
            let node = &self.nodes[index];
            node.word && rank_of(node.local.as_slice()).is_some_and(|value| value > rank)
        })
    }

    fn find_div(&self, xml: &[u8], parent: usize, id: Id) -> Result<Option<usize>> {
        for &index in &self.nodes[parent].children {
            let node = &self.nodes[index];
            if node.word && node.local.as_slice() == b"div" {
                if attribute_value(xml, node, b"id")?
                    .as_deref()
                    .is_some_and(|value| value.trim().parse::<i64>().ok() == Some(id.get()))
                {
                    return Ok(Some(index));
                }
                if let Some(found) = self.find_div(xml, index, id)? {
                    return Ok(Some(found));
                }
            } else if let Some(found) = self.find_div(xml, index, id)? {
                return Ok(Some(found));
            }
        }
        Ok(None)
    }
}

fn push_node(
    nodes: &mut Vec<Node>,
    start: usize,
    open_end: usize,
    parent: Option<usize>,
    element: BytesStart<'_>,
    word: bool,
    empty: bool,
) -> Result<usize> {
    if nodes.len() >= MAX_XML_EVENTS {
        return Err(invalid(format!(
            "web-settings XML exceeds {MAX_XML_EVENTS} elements"
        )));
    }
    let index = nodes.len();
    nodes.push(Node {
        start,
        open_end,
        close_start: open_end,
        end: open_end,
        children: Vec::new(),
        local: element.local_name().as_ref().to_vec(),
        qname: element.name().as_ref().to_vec(),
        word,
        empty,
    });
    if let Some(parent) = parent {
        nodes[parent]
            .children
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "web-settings rewrite child index",
                source,
            })?;
        nodes[parent].children.push(index);
    }
    Ok(index)
}

pub(crate) fn rewrite(source: &[u8], edits: &[Edit]) -> Result<Vec<u8>> {
    let mut xml = source.to_vec();
    for edit in edits {
        xml = rewrite_one(&xml, edit)?;
    }
    Ok(xml)
}

fn rewrite_one(source: &[u8], edit: &Edit) -> Result<Vec<u8>> {
    let tree = Tree::parse(source)?;
    let prefix = prefix_of(&tree.nodes[tree.root].qname);
    match edit {
        Edit::TargetScreen(value) => rewrite_root_leaf(
            &tree,
            source,
            b"targetScreenSz",
            11,
            value.map(Screen::as_str),
            &prefix,
        ),
        Edit::FramesetSize(value) => {
            rewrite_frameset_leaf(&tree, source, b"sz", 0, value.as_deref(), &prefix)
        },
        Edit::FramesetLayout(value) => rewrite_frameset_leaf(
            &tree,
            source,
            b"frameLayout",
            2,
            value.map(Layout::as_str),
            &prefix,
        ),
        Edit::DivBorder { id, side, value } => {
            rewrite_div_border(&tree, source, *id, *side, value.as_ref(), &prefix)
        },
    }
}

fn rewrite_root_leaf(
    tree: &Tree,
    source: &[u8],
    name: &[u8],
    rank: u8,
    value: Option<&str>,
    prefix: &str,
) -> Result<Vec<u8>> {
    let existing = tree.child(tree.root, name)?;
    match (existing, value) {
        (Some(index), Some(value)) => apply_range(
            source,
            tree.nodes[index].start,
            tree.nodes[index].end,
            value_element(prefix, name, value).into_bytes(),
        ),
        (Some(index), None) => apply_range(
            source,
            tree.nodes[index].start,
            tree.nodes[index].end,
            Vec::new(),
        ),
        (None, Some(value)) => insert_child(
            tree,
            source,
            tree.root,
            rank,
            value_element(prefix, name, value).into_bytes(),
        ),
        (None, None) => Ok(source.to_vec()),
    }
}

fn rewrite_frameset_leaf(
    tree: &Tree,
    source: &[u8],
    name: &[u8],
    rank: u8,
    value: Option<&str>,
    prefix: &str,
) -> Result<Vec<u8>> {
    let frameset = tree.child(tree.root, b"frameset")?;
    let Some(frameset) = frameset else {
        return match value {
            Some(value) => insert_child(
                tree,
                source,
                tree.root,
                0,
                frameset_element(prefix, name, value).into_bytes(),
            ),
            None => Ok(source.to_vec()),
        };
    };
    let existing = tree.child(frameset, name)?;
    match (existing, value) {
        (Some(index), Some(value)) => apply_range(
            source,
            tree.nodes[index].start,
            tree.nodes[index].end,
            value_element(prefix, name, value).into_bytes(),
        ),
        (Some(index), None) => apply_range(
            source,
            tree.nodes[index].start,
            tree.nodes[index].end,
            Vec::new(),
        ),
        (None, Some(value)) => insert_child(
            tree,
            source,
            frameset,
            rank,
            value_element(prefix, name, value).into_bytes(),
        ),
        (None, None) => Ok(source.to_vec()),
    }
}

fn rewrite_div_border(
    tree: &Tree,
    source: &[u8],
    id: Id,
    side: BorderSide,
    value: Option<&Border>,
    prefix: &str,
) -> Result<Vec<u8>> {
    let divs = tree
        .child(tree.root, b"divs")?
        .ok_or_else(|| invalid("web-settings division container is absent"))?;
    let div = tree
        .find_div(source, divs, id)?
        .ok_or_else(|| invalid(format!("HTML division '{id}' does not exist")))?;
    let borders = tree.child(div, b"divBdr")?;
    let Some(borders) = borders else {
        return match value {
            Some(value) => insert_child(
                tree,
                source,
                div,
                6,
                div_borders_element(prefix, side, value).into_bytes(),
            ),
            None => Ok(source.to_vec()),
        };
    };
    let side_name = side.as_str().as_bytes();
    let existing = tree.child(borders, side_name)?;
    match (existing, value) {
        (Some(index), Some(value)) => apply_range(
            source,
            tree.nodes[index].start,
            tree.nodes[index].end,
            border_element(prefix, side, value).into_bytes(),
        ),
        (Some(index), None) => apply_range(
            source,
            tree.nodes[index].start,
            tree.nodes[index].end,
            Vec::new(),
        ),
        (None, Some(value)) => insert_child(
            tree,
            source,
            borders,
            border_rank(side),
            border_element(prefix, side, value).into_bytes(),
        ),
        (None, None) => Ok(source.to_vec()),
    }
}

fn insert_child(
    tree: &Tree,
    source: &[u8],
    parent: usize,
    rank: u8,
    child: Vec<u8>,
) -> Result<Vec<u8>> {
    if tree.nodes[parent].empty {
        let node = &tree.nodes[parent];
        let raw = &source[node.start..node.end];
        if !raw.ends_with(b"/>") {
            return Err(invalid("web-settings empty-element range is malformed"));
        }
        let mut replacement = Vec::new();
        replacement
            .try_reserve(raw.len() + child.len() + node.qname.len() + 3)
            .map_err(|source| Error::Allocation {
                resource: "web-settings expanded element",
                source,
            })?;
        replacement.extend_from_slice(&raw[..raw.len() - 2]);
        replacement.push(b'>');
        replacement.extend_from_slice(&child);
        replacement.extend_from_slice(b"</");
        replacement.extend_from_slice(&node.qname);
        replacement.push(b'>');
        return apply_range(source, node.start, node.end, replacement);
    }

    let offset = tree.child_at_rank(parent, rank).map_or_else(
        || tree.nodes[parent].close_start,
        |index| tree.nodes[index].start,
    );
    apply_range(source, offset, offset, child)
}

fn apply_range(source: &[u8], start: usize, end: usize, replacement: Vec<u8>) -> Result<Vec<u8>> {
    if start > end || end > source.len() {
        return Err(invalid("web-settings rewrite range is outside the source"));
    }
    let new_len = source
        .len()
        .checked_sub(end - start)
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or_else(|| invalid("web-settings rewrite size overflow"))?;
    if new_len > MAX_XML_BYTES {
        return Err(invalid(format!(
            "web-settings XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(new_len)
        .map_err(|source| Error::Allocation {
            resource: "web-settings rewritten XML",
            source,
        })?;
    output.extend_from_slice(&source[..start]);
    output.extend_from_slice(&replacement);
    output.extend_from_slice(&source[end..]);
    Ok(output)
}

fn prefix_of(qname: &[u8]) -> String {
    qname
        .iter()
        .position(|byte| *byte == b':')
        .map_or_else(String::new, |position| {
            String::from_utf8_lossy(&qname[..position]).into_owned()
        })
}

fn qualified(prefix: &str, local: &[u8]) -> String {
    let local = String::from_utf8_lossy(local);
    if prefix.is_empty() {
        local.into_owned()
    } else {
        format!("{prefix}:{local}")
    }
}

fn value_element(prefix: &str, name: &[u8], value: &str) -> String {
    let element = qualified(prefix, name);
    let attribute = qualified(prefix, b"val");
    format!("<{element} {attribute}=\"{}\"/>", escape(value))
}

fn frameset_element(prefix: &str, name: &[u8], value: &str) -> String {
    let frameset = qualified(prefix, b"frameset");
    format!(
        "<{frameset}>{}</{frameset}>",
        value_element(prefix, name, value)
    )
}

fn div_borders_element(prefix: &str, side: BorderSide, value: &Border) -> String {
    let borders = qualified(prefix, b"divBdr");
    format!(
        "<{borders}>{}</{borders}>",
        border_element(prefix, side, value)
    )
}

fn border_element(prefix: &str, side: BorderSide, value: &Border) -> String {
    let element = qualified(prefix, side.as_str().as_bytes());
    let val = qualified(prefix, b"val");
    let mut xml = format!("<{element} {val}=\"{}\"", escape(value.style()));
    append_optional_attribute(&mut xml, prefix, b"color", value.color());
    append_optional_attribute(
        &mut xml,
        prefix,
        b"themeColor",
        value.theme_color().map(|value| value.as_str()),
    );
    append_optional_attribute(
        &mut xml,
        prefix,
        b"themeTint",
        value.theme_tint().map(|value| format!("{value:02X}")),
    );
    append_optional_attribute(
        &mut xml,
        prefix,
        b"themeShade",
        value.theme_shade().map(|value| format!("{value:02X}")),
    );
    append_optional_attribute(
        &mut xml,
        prefix,
        b"sz",
        value.size_eighth_points().map(|value| value.to_string()),
    );
    append_optional_attribute(
        &mut xml,
        prefix,
        b"space",
        value.space_points().map(|value| value.to_string()),
    );
    append_optional_attribute(
        &mut xml,
        prefix,
        b"shadow",
        value.shadow().map(|value| value.to_string()),
    );
    append_optional_attribute(
        &mut xml,
        prefix,
        b"frame",
        value.frame().map(|value| value.to_string()),
    );
    xml.push_str("/>");
    xml
}

fn append_optional_attribute<T: std::fmt::Display>(
    xml: &mut String,
    prefix: &str,
    name: &[u8],
    value: Option<T>,
) {
    if let Some(value) = value {
        let name = qualified(prefix, name);
        let _ = write!(xml, " {name}=\"{}\"", escape(&value.to_string()));
    }
}

fn attribute_value(xml: &[u8], node: &Node, name: &[u8]) -> Result<Option<String>> {
    let raw = &xml[node.start..node.open_end];
    if raw.len() < 3 || !raw.starts_with(b"<") || !raw.ends_with(b">") {
        return Err(invalid("web-settings element start range is malformed"));
    }
    let mut content = raw[1..raw.len() - 1].to_vec();
    if content.last() == Some(&b'/') {
        content.pop();
    }
    while content.last().is_some_and(u8::is_ascii_whitespace) {
        content.pop();
    }
    let name_len = content
        .iter()
        .position(u8::is_ascii_whitespace)
        .unwrap_or(content.len());
    let content = String::from_utf8(content)
        .map_err(|_| invalid("web-settings attribute range is not UTF-8"))?;
    let element = BytesStart::from_content(content, name_len);
    let mut result = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        if result.is_some() {
            return Err(invalid(
                "web-settings element has duplicate target attributes",
            ));
        }
        result = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, element.decoder())
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(result)
}

fn rank_of(name: &[u8]) -> Option<u8> {
    match name {
        b"frameset" => Some(0),
        b"divs" => Some(1),
        b"encoding" => Some(2),
        b"optimizeForBrowser" => Some(3),
        b"relyOnVML" => Some(4),
        b"allowPNG" => Some(5),
        b"doNotRelyOnCSS" => Some(6),
        b"doNotSaveAsSingleFile" => Some(7),
        b"doNotOrganizeInFolder" => Some(8),
        b"doNotUseLongFileNames" => Some(9),
        b"pixelsPerInch" => Some(10),
        b"targetScreenSz" => Some(11),
        b"saveSmartTagsAsXml" => Some(12),
        _ => None,
    }
}

fn border_rank(side: BorderSide) -> u8 {
    match side {
        BorderSide::Top => 0,
        BorderSide::Left => 1,
        BorderSide::Bottom => 2,
        BorderSide::Right => 3,
    }
}
