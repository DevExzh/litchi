//! Structure-preserving mixed text content for spreadsheet cells.

use super::hyperlink::Link;
use super::{
    AnnotationElement, AnnotationNode,
    annotation::{decode_reference, parse_element, standard_namespace_uri},
};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    Decoder, XmlVersion,
    events::{BytesCData, BytesRef, BytesStart, BytesText},
};
use std::{
    borrow::Cow,
    collections::{BTreeMap, BTreeSet},
    ops::Range,
};

const MAX_RICH_TEXT_NODES: usize = 65_536;
const MAX_RICH_TEXT_DEPTH: usize = 128;
const MAX_RICH_TEXT_BYTES: usize = 16 * 1024 * 1024;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

/// Ordered, namespace-aware ODF paragraph content retained by an ODS cell.
///
/// The tree preserves spans, whitespace elements, fields, and extension
/// elements when a parsed spreadsheet is saved. Hyperlinks remain inert.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CellTextContent {
    paragraphs: Vec<AnnotationElement>,
    namespaces: BTreeMap<String, String>,
}

impl CellTextContent {
    /// Return the retained top-level `text:p` elements.
    pub fn paragraphs(&self) -> &[AnnotationElement] {
        &self.paragraphs
    }

    /// Return namespace bindings needed by retained extension content.
    pub fn namespaces(&self) -> &BTreeMap<String, String> {
        &self.namespaces
    }

    /// Return displayed text, separating adjacent paragraphs with a newline.
    pub fn plain_text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|paragraph| element_plain_text(paragraph, &self.namespaces))
            .collect::<Vec<_>>()
            .join("\n")
    }

    pub(crate) fn from_hyperlink(hyperlink: &Link) -> Result<Self> {
        let mut paragraph = AnnotationElement::new("text:p")?;
        paragraph
            .children
            .push(AnnotationNode::Element(hyperlink_element(
                hyperlink,
                vec![AnnotationNode::Text(hyperlink.text.clone())],
            )?));
        Ok(Self {
            paragraphs: vec![paragraph],
            namespaces: BTreeMap::new(),
        })
    }

    pub(crate) fn hyperlink_count(&self) -> usize {
        fn count(nodes: &[AnnotationNode], namespaces: &BTreeMap<String, String>) -> usize {
            nodes
                .iter()
                .map(|node| match node {
                    AnnotationNode::Text(_) => 0,
                    AnnotationNode::Element(element) => {
                        let scoped = scoped_namespaces(element, namespaces);
                        usize::from(is_text_anchor(element, &scoped))
                            + count(&element.children, &scoped)
                    },
                })
                .sum()
        }
        self.paragraphs
            .iter()
            .map(|paragraph| count(&paragraph.children, &self.namespaces))
            .sum()
    }

    pub(crate) fn wrap_hyperlink(
        &mut self,
        range: Range<usize>,
        hyperlink: &Link,
    ) -> Result<()> {
        let mut offset = 0usize;
        for paragraph in &mut self.paragraphs {
            let length = element_plain_text(paragraph, &self.namespaces).len();
            let end = offset
                .checked_add(length)
                .ok_or_else(|| invalid("cell rich-text offset overflow"))?;
            if range.start >= offset && range.end <= end {
                let local = range.start - offset..range.end - offset;
                let (before, remainder) =
                    split_nodes_at(&paragraph.children, local.start, &self.namespaces)?;
                let (selected, after) =
                    split_nodes_at(&remainder, local.end - local.start, &self.namespaces)?;
                let mut children = before;
                children.push(AnnotationNode::Element(hyperlink_element(
                    hyperlink, selected,
                )?));
                children.extend(after);
                paragraph.children = children;
                return Ok(());
            }
            offset = end
                .checked_add(1)
                .ok_or_else(|| invalid("cell rich-text offset overflow"))?;
        }
        Err(invalid(
            "cell hyperlink range crosses a paragraph boundary or is out of bounds",
        ))
    }

    pub(crate) fn remove_hyperlink(&mut self, index: usize) -> bool {
        fn remove(
            nodes: &mut Vec<AnnotationNode>,
            target: usize,
            seen: &mut usize,
            namespaces: &BTreeMap<String, String>,
        ) -> bool {
            let mut index_in_parent = 0usize;
            while index_in_parent < nodes.len() {
                let is_anchor = matches!(
                    &nodes[index_in_parent],
                    AnnotationNode::Element(element)
                        if is_text_anchor(element, &scoped_namespaces(element, namespaces))
                );
                if is_anchor {
                    if *seen == target {
                        let AnnotationNode::Element(element) = nodes.remove(index_in_parent) else {
                            unreachable!("checked element node");
                        };
                        nodes.splice(index_in_parent..index_in_parent, element.children);
                        return true;
                    }
                    *seen += 1;
                }
                if let AnnotationNode::Element(element) = &mut nodes[index_in_parent] {
                    let scoped = scoped_namespaces(element, namespaces);
                    if remove(&mut element.children, target, seen, &scoped) {
                        return true;
                    }
                }
                index_in_parent += 1;
            }
            false
        }

        let mut seen = 0usize;
        self.paragraphs
            .iter_mut()
            .any(|paragraph| remove(&mut paragraph.children, index, &mut seen, &self.namespaces))
    }

    pub(crate) fn clear_hyperlinks(&mut self) {
        fn unwrap(nodes: &mut Vec<AnnotationNode>, namespaces: &BTreeMap<String, String>) {
            let mut output = Vec::with_capacity(nodes.len());
            for node in std::mem::take(nodes) {
                match node {
                    AnnotationNode::Element(mut element) => {
                        let scoped = scoped_namespaces(&element, namespaces);
                        let anchor = is_text_anchor(&element, &scoped);
                        unwrap(&mut element.children, &scoped);
                        if anchor {
                            output.extend(element.children);
                        } else {
                            output.push(AnnotationNode::Element(element));
                        }
                    },
                    AnnotationNode::Text(text) => output.push(AnnotationNode::Text(text)),
                }
            }
            *nodes = output;
        }
        for paragraph in &mut self.paragraphs {
            unwrap(&mut paragraph.children, &self.namespaces);
        }
    }

    pub(crate) fn synchronize_hyperlinks(&mut self, hyperlinks: &[Link]) -> bool {
        fn synchronize(
            nodes: &mut [AnnotationNode],
            hyperlinks: &[Link],
            index: &mut usize,
            namespaces: &BTreeMap<String, String>,
        ) -> bool {
            for node in nodes {
                let AnnotationNode::Element(element) = node else {
                    continue;
                };
                let scoped = scoped_namespaces(element, namespaces);
                if is_text_anchor(element, &scoped) {
                    let Some(hyperlink) = hyperlinks.get(*index) else {
                        return false;
                    };
                    let name = element.name.clone();
                    let attributes = std::mem::take(&mut element.attributes);
                    let children = std::mem::take(&mut element.children);
                    let Ok(mut replacement) = hyperlink_element(hyperlink, children) else {
                        return false;
                    };
                    replacement.name = name;
                    for (name, value) in attributes {
                        if is_namespace_declaration(&name)
                            || !is_managed_hyperlink_attribute(&name, &scoped)
                        {
                            replacement.attributes.entry(name).or_insert(value);
                        }
                    }
                    *element = replacement;
                    *index += 1;
                } else if !synchronize(&mut element.children, hyperlinks, index, &scoped) {
                    return false;
                }
            }
            true
        }

        let mut index = 0usize;
        let complete = self.paragraphs.iter_mut().all(|paragraph| {
            synchronize(
                &mut paragraph.children,
                hyperlinks,
                &mut index,
                &self.namespaces,
            )
        });
        complete && index == hyperlinks.len()
    }

    pub(crate) fn write_xml(&self, output: &mut String) {
        for paragraph in &self.paragraphs {
            write_top_level_element(output, paragraph, &self.namespaces);
        }
    }
}

fn is_text_anchor(element: &AnnotationElement, namespaces: &BTreeMap<String, String>) -> bool {
    is_text_element(element, "a", namespaces)
}

fn is_text_element(
    element: &AnnotationElement,
    expected_local: &str,
    namespaces: &BTreeMap<String, String>,
) -> bool {
    let (prefix, local) = element.name.rsplit_once(':').unwrap_or(("", &element.name));
    local == expected_local
        && namespace_uri(element, prefix, namespaces)
            == Some("urn:oasis:names:tc:opendocument:xmlns:text:1.0")
}

fn namespace_uri<'a>(
    element: &'a AnnotationElement,
    prefix: &str,
    namespaces: &'a BTreeMap<String, String>,
) -> Option<&'a str> {
    let declaration = if prefix.is_empty() {
        "xmlns".to_string()
    } else {
        format!("xmlns:{prefix}")
    };
    element
        .attributes
        .get(&declaration)
        .or_else(|| namespaces.get(prefix))
        .map(String::as_str)
        .or_else(|| standard_namespace_uri(prefix))
}

fn scoped_namespaces<'a>(
    element: &AnnotationElement,
    namespaces: &'a BTreeMap<String, String>,
) -> Cow<'a, BTreeMap<String, String>> {
    let declarations = element.attributes.iter().filter_map(|(name, value)| {
        if name == "xmlns" {
            Some(("", value))
        } else {
            name.strip_prefix("xmlns:").map(|prefix| (prefix, value))
        }
    });
    let mut scoped = None;
    for (prefix, value) in declarations {
        scoped
            .get_or_insert_with(|| namespaces.clone())
            .insert(prefix.to_string(), value.clone());
    }
    scoped.map_or(Cow::Borrowed(namespaces), Cow::Owned)
}

fn is_namespace_declaration(name: &str) -> bool {
    name == "xmlns" || name.starts_with("xmlns:")
}

fn is_managed_hyperlink_attribute(name: &str, namespaces: &BTreeMap<String, String>) -> bool {
    let Some((prefix, local)) = name.rsplit_once(':') else {
        return false;
    };
    let namespace = namespaces
        .get(prefix)
        .map(String::as_str)
        .or_else(|| standard_namespace_uri(prefix));
    matches!(
        (namespace, local),
        (
            Some("http://www.w3.org/1999/xlink"),
            "type" | "href" | "actuate" | "show"
        ) | (
            Some("urn:oasis:names:tc:opendocument:xmlns:office:1.0"),
            "target-frame-name" | "name" | "title"
        ) | (
            Some("urn:oasis:names:tc:opendocument:xmlns:text:1.0"),
            "style-name" | "visited-style-name"
        )
    )
}

fn element_plain_text(
    element: &AnnotationElement,
    namespaces: &BTreeMap<String, String>,
) -> String {
    let scoped = scoped_namespaces(element, namespaces);
    if is_text_element(element, "s", &scoped) {
        let count = element
            .attributes
            .iter()
            .find_map(|(name, value)| {
                let (prefix, local) = name.rsplit_once(':')?;
                (local == "c"
                    && namespace_uri(element, prefix, &scoped)
                        == Some("urn:oasis:names:tc:opendocument:xmlns:text:1.0"))
                .then_some(value)
            })
            .and_then(|value| value.parse::<usize>().ok())
            .unwrap_or(1)
            .min(1_000_000);
        return " ".repeat(count);
    }
    if is_text_element(element, "tab", &scoped) {
        return "\t".to_string();
    }
    if is_text_element(element, "line-break", &scoped) {
        return "\n".to_string();
    }

    let mut text = String::new();
    for child in &element.children {
        match child {
            AnnotationNode::Text(value) => text.push_str(value),
            AnnotationNode::Element(element) => {
                text.push_str(&element_plain_text(element, &scoped));
            },
        }
    }
    text
}

fn hyperlink_element(
    hyperlink: &Link,
    children: Vec<AnnotationNode>,
) -> Result<AnnotationElement> {
    let mut element = AnnotationElement::new("text:a")?;
    element.set_attribute("xlink:type", "simple")?;
    element.set_attribute("xlink:href", hyperlink.href.clone())?;
    for (name, value) in [
        (
            "xlink:actuate",
            hyperlink.actuate.map(|value| value.as_str().to_string()),
        ),
        (
            "office:target-frame-name",
            hyperlink.target_frame_name.clone(),
        ),
        (
            "xlink:show",
            hyperlink.show.map(|value| value.as_str().to_string()),
        ),
        ("office:name", hyperlink.name.clone()),
        ("office:title", hyperlink.title.clone()),
        ("text:style-name", hyperlink.style_name.clone()),
        (
            "text:visited-style-name",
            hyperlink.visited_style_name.clone(),
        ),
    ] {
        if let Some(value) = value {
            element.set_attribute(name, value)?;
        }
    }
    element.children = children;
    Ok(element)
}

fn split_nodes_at(
    nodes: &[AnnotationNode],
    offset: usize,
    namespaces: &BTreeMap<String, String>,
) -> Result<(Vec<AnnotationNode>, Vec<AnnotationNode>)> {
    let total = nodes
        .iter()
        .try_fold(0usize, |sum, node| {
            sum.checked_add(node_plain_len(node, namespaces))
        })
        .ok_or_else(|| invalid("cell rich-text size overflow"))?;
    if offset > total {
        return Err(invalid("cell rich-text split offset is out of bounds"));
    }

    let mut left = Vec::new();
    let mut right = Vec::new();
    let mut cursor = 0usize;
    let mut split = false;
    for node in nodes {
        if split {
            right.push(node.clone());
            continue;
        }
        let length = node_plain_len(node, namespaces);
        let end = cursor
            .checked_add(length)
            .ok_or_else(|| invalid("cell rich-text size overflow"))?;
        if offset >= end {
            left.push(node.clone());
            cursor = end;
            continue;
        }
        let local = offset - cursor;
        let (node_left, node_right) = split_node_at(node, local, namespaces)?;
        left.extend(node_left);
        right.extend(node_right);
        split = true;
    }
    Ok((left, right))
}

fn split_node_at(
    node: &AnnotationNode,
    offset: usize,
    namespaces: &BTreeMap<String, String>,
) -> Result<(Option<AnnotationNode>, Option<AnnotationNode>)> {
    let length = node_plain_len(node, namespaces);
    if offset == 0 {
        return Ok((None, Some(node.clone())));
    }
    if offset == length {
        return Ok((Some(node.clone()), None));
    }
    match node {
        AnnotationNode::Text(text) => {
            if !text.is_char_boundary(offset) {
                return Err(invalid(
                    "cell hyperlink range is not on a UTF-8 character boundary",
                ));
            }
            Ok((
                Some(AnnotationNode::Text(text[..offset].to_string())),
                Some(AnnotationNode::Text(text[offset..].to_string())),
            ))
        },
        AnnotationNode::Element(element) => {
            let scoped = scoped_namespaces(element, namespaces);
            if is_text_anchor(element, &scoped) {
                return Err(invalid(
                    "cell hyperlink range overlaps an existing hyperlink",
                ));
            }
            if element.children.is_empty() {
                return Err(invalid(
                    "cell hyperlink range splits an atomic inline element",
                ));
            }
            let (left, right) = split_nodes_at(&element.children, offset, &scoped)?;
            let mut left_element = element.clone();
            left_element.children = left;
            let mut right_element = element.clone();
            right_element.children = right;
            Ok((
                Some(AnnotationNode::Element(left_element)),
                Some(AnnotationNode::Element(right_element)),
            ))
        },
    }
}

fn node_plain_len(node: &AnnotationNode, namespaces: &BTreeMap<String, String>) -> usize {
    match node {
        AnnotationNode::Text(text) => text.len(),
        AnnotationNode::Element(element) => element_plain_text(element, namespaces).len(),
    }
}

fn write_top_level_element(
    output: &mut String,
    element: &AnnotationElement,
    namespaces: &BTreeMap<String, String>,
) {
    output.push('<');
    output.push_str(&element.name);
    for (name, value) in &element.attributes {
        write_attribute(output, name, value);
    }
    let mut used = BTreeSet::new();
    element.collect_prefixes(&mut used);
    let mut declared = BTreeSet::new();
    element.collect_namespace_declarations(&mut declared);
    for prefix in used {
        if standard_namespace_uri(&prefix).is_none()
            && !declared.contains(&prefix)
            && let Some(uri) = namespaces.get(&prefix)
        {
            write_attribute(output, &format!("xmlns:{prefix}"), uri);
        }
    }
    if element.children.is_empty() {
        output.push_str("/>");
        return;
    }
    output.push('>');
    for child in &element.children {
        match child {
            AnnotationNode::Text(text) => output.push_str(&escape_xml(text)),
            AnnotationNode::Element(element) => element.write_xml(output),
        }
    }
    output.push_str("</");
    output.push_str(&element.name);
    output.push('>');
}

fn write_attribute(output: &mut String, name: &str, value: &str) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    output.push_str(&escape_xml(value));
    output.push('"');
}

pub(crate) struct CellTextContentBuilder {
    content: CellTextContent,
    stack: Vec<AnnotationElement>,
    nodes: usize,
    text_bytes: usize,
}

impl CellTextContentBuilder {
    pub(crate) fn new(namespaces: BTreeMap<String, String>) -> Self {
        Self {
            content: CellTextContent {
                paragraphs: Vec::new(),
                namespaces,
            },
            stack: Vec::new(),
            nodes: 0,
            text_bytes: 0,
        }
    }

    pub(crate) fn start(&mut self, start: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if self.stack.len() >= MAX_RICH_TEXT_DEPTH {
            return Err(invalid("cell rich text exceeds the XML depth limit"));
        }
        self.add_node()?;
        self.stack.push(parse_element(start, decoder)?);
        Ok(())
    }

    pub(crate) fn empty(&mut self, start: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        self.add_node()?;
        self.add_element(parse_element(start, decoder)?)
    }

    pub(crate) fn text(&mut self, text: &BytesText<'_>) -> Result<()> {
        let text = text
            .xml_content(XmlVersion::Explicit1_0)
            .map_err(|error| invalid(format!("invalid cell rich text: {error}")))?
            .into_owned();
        self.push_text(text)
    }

    pub(crate) fn cdata(&mut self, text: &BytesCData<'_>) -> Result<()> {
        let text = text
            .xml_content(XmlVersion::Explicit1_0)
            .map_err(|error| invalid(format!("invalid cell rich CDATA: {error}")))?
            .into_owned();
        self.push_text(text)
    }

    pub(crate) fn reference(&mut self, reference: &BytesRef<'_>) -> Result<()> {
        self.push_text(decode_reference(reference)?)
    }

    pub(crate) fn end(&mut self) -> Result<()> {
        let element = self
            .stack
            .pop()
            .ok_or_else(|| invalid("unbalanced cell rich-text element"))?;
        self.add_element(element)
    }

    pub(crate) fn finish(self) -> Result<CellTextContent> {
        if !self.stack.is_empty() {
            return Err(invalid("unclosed cell rich-text element"));
        }
        Ok(self.content)
    }

    fn push_text(&mut self, text: String) -> Result<()> {
        self.text_bytes = self
            .text_bytes
            .checked_add(text.len())
            .ok_or_else(|| invalid("cell rich-text size overflow"))?;
        if self.text_bytes > MAX_RICH_TEXT_BYTES {
            return Err(invalid("cell rich text exceeds the text-size limit"));
        }
        if let Some(parent) = self.stack.last_mut() {
            parent.children.push(AnnotationNode::Text(text));
        }
        Ok(())
    }

    fn add_node(&mut self) -> Result<()> {
        self.nodes = self
            .nodes
            .checked_add(1)
            .ok_or_else(|| invalid("cell rich-text node count overflow"))?;
        if self.nodes > MAX_RICH_TEXT_NODES {
            return Err(invalid("cell rich text exceeds the node limit"));
        }
        Ok(())
    }

    fn add_element(&mut self, element: AnnotationElement) -> Result<()> {
        if let Some(parent) = self.stack.last_mut() {
            parent.children.push(AnnotationNode::Element(element));
        } else {
            self.content.paragraphs.push(element);
        }
        Ok(())
    }
}
