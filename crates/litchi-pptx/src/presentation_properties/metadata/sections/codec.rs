//! Typed, bounded support for `PowerPoint` 2010 presentation sections.

use super::super::{BoundedString, Limits, process_ooxml};
use super::model::{List, Section};
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
use std::collections::HashSet;

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const SECTION_URI: &str = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
impl Section {
    /// Serialize this `p14:section` element.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn to_xml(&self, limits: Limits) -> Result<String> {
        validate_section(self, limits)?;
        let mut xml = BoundedString::new(limits, "presentation section output");
        self.write_xml(&mut xml)?;
        Ok(xml.finish())
    }

    fn write_xml(&self, xml: &mut BoundedString) -> Result<()> {
        xml.push_str("<p14:section")?;
        if let Some(name) = &self.name {
            xml.push_str(" name=\"")?;
            xml.push_escaped_xml(name)?;
            xml.push_str("\"")?;
        }
        if let Some(id) = &self.id {
            xml.push_str(" id=\"")?;
            xml.push_escaped_xml(id)?;
            xml.push_str("\"")?;
        }
        xml.push_str("><p14:sldIdLst>")?;
        for slide_id in &self.slide_ids {
            xml.write_fmt(format_args!("<p14:sldId id=\"{slide_id}\"/>"))?;
        }
        xml.push_str("</p14:sldIdLst>")?;
        if let Some(extension) = &self.extension_xml {
            xml.push_str(std::str::from_utf8(extension).map_err(xml_error)?)?;
        }
        xml.push_str("</p14:section>")?;
        Ok(())
    }
}

impl List {
    /// Replace a section while retaining its stable GUID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_by_id(
        &mut self,
        id: &str,
        mut replacement: Section,
        limits: Limits,
    ) -> Result<()> {
        let target = self
            .get_by_id_mut(id)
            .ok_or_else(|| invalid(format!("section {id} was not found")))?;
        replacement.id = Some(id.to_owned());
        *target = replacement;
        validate_list(self, limits)
    }

    /// Parse the section extension from a complete `p:presentation` document.
    ///
    /// Unrelated presentation extensions remain inert and are never evaluated.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_xml(xml: &[u8], limits: Limits) -> Result<Self> {
        if xml.len() > limits.max_input_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::InputBytes,
                xml.len(),
                limits.max_input_bytes(),
            ));
        }
        let processed = process_ooxml(xml, limits)?;
        if processed.len() > limits.max_output_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::OutputBytes,
                processed.len(),
                limits.max_output_bytes(),
            ));
        }
        Self::from_processed(processed.as_ref(), limits)
    }

    /// Parse XML that has already crossed the caller's MCE/resource boundary.
    ///
    /// Source-backed owners use this entry point after applying their
    /// effective physical and semantic limits. It deliberately performs no
    /// second MCE pass, while retaining the section codec's own bounded DOM
    /// and structural validation.
    pub(crate) fn from_processed(xml: &[u8], limits: Limits) -> Result<Self> {
        if xml.len() > limits.max_output_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::OutputBytes,
                xml.len(),
                limits.max_output_bytes(),
            ));
        }
        parse_sections(&parse_dom(xml, limits)?, limits)
    }

    /// Serialize the complete `p:extLst` fragment containing `p14:sectionLst`.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn to_xml(&self, limits: Limits) -> Result<String> {
        if self.sections.is_empty() {
            return Ok(String::new());
        }
        validate_list(self, limits)?;
        let mut xml = BoundedString::new(limits, "serialized presentation sections");
        xml.push_str("<p:extLst><p:ext uri=\"")?;
        xml.push_str(SECTION_URI)?;
        xml.push_str("\"><p14:sectionLst xmlns:p14=\"")?;
        xml.push_str(P14)?;
        xml.push_str("\">")?;
        for section in &self.sections {
            section.write_xml(&mut xml)?;
        }
        xml.push_str("</p14:sectionLst></p:ext></p:extLst>")?;
        Ok(xml.finish())
    }
}

#[derive(Clone)]
struct Attr {
    qname: String,
    namespace: String,
    local: String,
    value: String,
}

#[derive(Clone)]
enum Content {
    Node(Node),
    Text(String),
    CData(String),
    Comment(String),
}

#[derive(Clone)]
struct Node {
    qname: String,
    namespace: String,
    local: String,
    attributes: Vec<Attr>,
    bindings: Vec<(String, String)>,
    content: Vec<Content>,
}

fn parse_dom(xml: &[u8], limits: Limits) -> Result<Node> {
    if xml.len() > limits.max_output_bytes() {
        return Err(limit(
            crate::opened::CodecLimitKind::OutputBytes,
            xml.len(),
            limits.max_output_bytes(),
        ));
    }
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut reader = Reader::from_reader(xml);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    loop {
        let decoder = reader.decoder();
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::opened::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        limits.max_nodes(),
                    )
                })?;
                if nodes > limits.max_nodes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::XmlNodes,
                        nodes,
                        limits.max_nodes(),
                    ));
                }
                if stack.len() >= limits.max_depth() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::XmlDepth,
                        stack.len() + 1,
                        limits.max_depth(),
                    ));
                }
                stack.push(make_node(&element, decoder, &stack, limits)?);
            },
            Ok(Event::Empty(element)) => {
                check_empty_depth(stack.len(), limits)?;
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::opened::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        limits.max_nodes(),
                    )
                })?;
                if nodes > limits.max_nodes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::XmlNodes,
                        nodes,
                        limits.max_nodes(),
                    ));
                }
                let node = make_node(&element, decoder, &stack, limits)?;
                attach(&mut stack, &mut root, node)?;
            },
            Ok(Event::End(_)) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element"))?;
                attach(&mut stack, &mut root, node)?;
            },
            Ok(Event::Text(text)) => {
                if text.as_ref().len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::TextBytes,
                        text.as_ref().len(),
                        limits.max_string_bytes(),
                    ));
                }
                let text = text.decode().map_err(xml_error)?;
                if text.len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::StringBytes,
                        text.len(),
                        limits.max_string_bytes(),
                    ));
                }
                let text = text.into_owned();
                if let Some(node) = stack.last_mut() {
                    node.content.push(Content::Text(text));
                } else if !text.trim().is_empty() {
                    return Err(invalid("text outside presentation root"));
                }
            },
            Ok(Event::CData(text)) => {
                if text.as_ref().len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::TextBytes,
                        text.as_ref().len(),
                        limits.max_string_bytes(),
                    ));
                }
                let text = text.decode().map_err(xml_error)?;
                if text.len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::StringBytes,
                        text.len(),
                        limits.max_string_bytes(),
                    ));
                }
                let text = text.into_owned();
                if let Some(node) = stack.last_mut() {
                    node.content.push(Content::CData(text));
                } else {
                    return Err(invalid("CDATA outside presentation root"));
                }
            },
            Ok(Event::Comment(text)) => {
                if text.as_ref().len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::TextBytes,
                        text.as_ref().len(),
                        limits.max_string_bytes(),
                    ));
                }
                let text = text.decode().map_err(xml_error)?;
                if text.len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::StringBytes,
                        text.len(),
                        limits.max_string_bytes(),
                    ));
                }
                let text = text.into_owned();
                if let Some(node) = stack.last_mut() {
                    node.content.push(Content::Comment(text));
                }
            },
            Ok(Event::GeneralRef(reference)) => {
                if reference.as_ref().len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::ReferenceBytes,
                        reference.as_ref().len(),
                        limits.max_string_bytes(),
                    ));
                }
                let text = litchi_ooxml_common::xml::decode_xml_reference(&reference)?;
                if text.len() > limits.max_string_bytes() {
                    return Err(limit(
                        crate::opened::CodecLimitKind::StringBytes,
                        text.len(),
                        limits.max_string_bytes(),
                    ));
                }
                if let Some(node) = stack.last_mut() {
                    node.content.push(Content::Text(text));
                } else {
                    return Err(invalid("entity outside presentation root"));
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Decl(_)) => {},
            Ok(Event::Eof) => break,
            Err(error) => return Err(xml_error(error)),
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated presentation-section XML"));
    }
    root.ok_or_else(|| invalid("missing presentation root"))
}

fn make_node(
    element: &BytesStart<'_>,
    decoder: Decoder,
    stack: &[Node],
    limits: Limits,
) -> Result<Node> {
    let raw_name = element.name();
    let qname = std::str::from_utf8(raw_name.as_ref()).map_err(xml_error)?;
    bounded_string(qname, limits)?;
    let qname = qname.to_owned();
    let mut bindings = stack
        .last()
        .map(|node| node.bindings.clone())
        .unwrap_or_default();
    let mut raw = Vec::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        if raw.len() >= limits.max_attributes() {
            return Err(limit(
                crate::opened::CodecLimitKind::XmlAttributes,
                raw.len() + 1,
                limits.max_attributes(),
            ));
        }
        let name = std::str::from_utf8(attribute.key.as_ref()).map_err(xml_error)?;
        bounded_string(name, limits)?;
        if attribute.value.as_ref().len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::XmlAttributeBytes,
                attribute.value.as_ref().len(),
                limits.max_string_bytes(),
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?;
        if value.len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::StringBytes,
                value.len(),
                limits.max_string_bytes(),
            ));
        }
        raw.push((name.to_owned(), value.into_owned()));
    }
    for (name, value) in &raw {
        if name == "xmlns" || name.starts_with("xmlns:") {
            let prefix = name.strip_prefix("xmlns:").unwrap_or("").to_owned();
            if let Some(binding) = bindings.iter_mut().find(|binding| binding.0 == prefix) {
                binding.1.clone_from(value);
            } else {
                bindings.push((prefix, value.clone()));
            }
        }
    }
    let (prefix, local) = split_qname(&qname)?;
    let namespace = resolve(&bindings, prefix)?;
    let local = local.to_owned();
    let mut attributes = Vec::new();
    for (name, value) in raw {
        if name == "xmlns" || name.starts_with("xmlns:") {
            continue;
        }
        let (prefix, local) = split_qname(&name)?;
        let namespace = if prefix.is_empty() {
            String::new()
        } else {
            resolve(&bindings, prefix)?
        };
        let local = local.to_owned();
        attributes.push(Attr {
            qname: name,
            namespace,
            local,
            value,
        });
    }
    Ok(Node {
        qname,
        namespace,
        local,
        attributes,
        bindings,
        content: Vec::new(),
    })
}

fn attach(stack: &mut [Node], root: &mut Option<Node>, node: Node) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.content.push(Content::Node(node));
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

fn parse_sections(root: &Node, limits: Limits) -> Result<List> {
    expect(root, &[P, PS], "presentation")?;
    let mut presentation_ext = None;
    for child in children(root)? {
        if is(child, &[P, PS], "extLst") {
            if child.namespace != root.namespace {
                return Err(invalid(
                    "presentation section extension has mixed conformance",
                ));
            }
            if presentation_ext.replace(child).is_some() {
                return Err(invalid("duplicate presentation extLst"));
            }
        }
    }
    let Some(extension_list) = presentation_ext else {
        return Ok(List::new());
    };
    let mut section_list = None;
    for extension in children(extension_list)? {
        expect(extension, &[root.namespace.as_str()], "ext")?;
        let uri = required_attr(extension, "uri")?;
        if uri == SECTION_URI {
            if section_list.is_some() {
                return Err(invalid("duplicate PowerPoint section extension"));
            }
            only_unqualified(extension, &["uri"])?;
            let payload = children(extension)?;
            if payload.len() != 1 {
                return Err(invalid("section extension requires one sectionLst payload"));
            }
            section_list = Some(parse_section_list(payload[0], limits)?);
        }
    }
    Ok(section_list.unwrap_or_default())
}

fn parse_section_list(node: &Node, limits: Limits) -> Result<List> {
    expect(node, &[P14], "sectionLst")?;
    no_attributes(node)?;
    let sections = children(node)?;
    if sections.is_empty() {
        return Err(invalid("sectionLst requires at least one section"));
    }
    if sections.len() > limits.max_sections() {
        return Err(limit(
            crate::opened::CodecLimitKind::Sections,
            sections.len(),
            limits.max_sections(),
        ));
    }
    let mut parsed = Vec::with_capacity(sections.len());
    let mut section_ids = HashSet::new();
    let mut total_slide_ids = 0usize;
    for node in sections {
        expect(node, &[P14], "section")?;
        let name = optional_attr(node, "name")?;
        let id = optional_attr(node, "id")?;
        only_unqualified(node, &["name", "id"])?;
        if let Some(name) = &name {
            bounded_string(name, limits)?;
        }
        if let Some(id) = &id {
            validate_guid(id)?;
            if !section_ids.insert(id.clone()) {
                return Err(invalid("duplicate section GUID"));
            }
        }
        let content = children(node)?;
        if content.is_empty()
            || content.len() > 2
            || !is(content[0], &[P14], "sldIdLst")
            || (content.len() == 2 && !is(content[1], &[P, PS], "extLst"))
        {
            return Err(invalid(
                "section requires sldIdLst followed by optional extLst",
            ));
        }
        let slide_ids = parse_slide_id_list(content[0], &mut total_slide_ids, limits)?;
        let extension_xml = if content.len() == 2 {
            validate_extension_list(content[1], limits)?;
            Some(node_xml(content[1], limits)?)
        } else {
            None
        };
        parsed.push(Section {
            name,
            id,
            slide_ids,
            extension_xml,
        });
    }
    Ok(List { sections: parsed })
}

fn parse_slide_id_list(node: &Node, total: &mut usize, limits: Limits) -> Result<Vec<u32>> {
    no_attributes(node)?;
    let mut slide_ids = Vec::new();
    let mut seen = HashSet::new();
    for slide in children(node)? {
        expect(slide, &[P14], "sldId")?;
        let id = required_attr(slide, "id")?
            .parse::<u32>()
            .map_err(|_err| invalid("invalid section slide ID"))?;
        if id < 256 {
            return Err(invalid("section slide ID is below 256"));
        }
        only_unqualified(slide, &["id"])?;
        leaf(slide)?;
        if !seen.insert(id) {
            return Err(invalid("duplicate slide ID within section"));
        }
        *total = total
            .checked_add(1)
            .ok_or_else(|| invalid("section slide reference count overflow"))?;
        if *total > limits.max_slide_references() {
            return Err(limit(
                crate::opened::CodecLimitKind::SlideReferences,
                *total,
                limits.max_slide_references(),
            ));
        }
        slide_ids.push(id);
    }
    Ok(slide_ids)
}

fn validate_extension_list(node: &Node, limits: Limits) -> Result<()> {
    expect(node, &[P, PS], "extLst")?;
    no_attributes(node)?;
    let extensions = children(node)?;
    if extensions.len() > limits.max_extensions() {
        return Err(limit(
            crate::opened::CodecLimitKind::Extensions,
            extensions.len(),
            limits.max_extensions(),
        ));
    }
    for extension in extensions {
        expect(extension, &[P, PS], "ext")?;
        let uri = required_attr(extension, "uri")?;
        if uri.is_empty() {
            return Err(invalid("section extension URI is empty"));
        }
        bounded_string(&uri, limits)?;
        only_unqualified(extension, &["uri"])?;
    }
    Ok(())
}

fn validate_list(list: &List, limits: Limits) -> Result<()> {
    if list.sections.len() > limits.max_sections() {
        return Err(limit(
            crate::opened::CodecLimitKind::Sections,
            list.sections.len(),
            limits.max_sections(),
        ));
    }
    let mut section_ids = HashSet::new();
    let mut total = 0usize;
    for section in &list.sections {
        validate_section(section, limits)?;
        if let Some(id) = &section.id
            && !section_ids.insert(id)
        {
            return Err(invalid("duplicate section GUID"));
        }
        total = total
            .checked_add(section.slide_ids.len())
            .ok_or_else(|| invalid("section slide reference count overflow"))?;
        if total > limits.max_slide_references() {
            return Err(limit(
                crate::opened::CodecLimitKind::SlideReferences,
                total,
                limits.max_slide_references(),
            ));
        }
    }
    Ok(())
}

fn validate_section(section: &Section, limits: Limits) -> Result<()> {
    if let Some(name) = &section.name {
        bounded_string(name, limits)?;
    }
    if let Some(id) = &section.id {
        validate_guid(id)?;
    }
    let mut slides = HashSet::new();
    for id in &section.slide_ids {
        if *id < 256 {
            return Err(invalid("section slide ID is below 256"));
        }
        if !slides.insert(*id) {
            return Err(invalid("duplicate slide ID within section"));
        }
    }
    if let Some(extension) = &section.extension_xml {
        if extension.len() > limits.max_string_bytes() {
            return Err(limit(
                crate::opened::CodecLimitKind::ExtensionBytes,
                extension.len(),
                limits.max_string_bytes(),
            ));
        }
        validate_extension_list(&parse_dom(extension, limits)?, limits)?;
    }
    Ok(())
}

fn validate_guid(value: &str) -> Result<()> {
    let bytes = value.as_bytes();
    let valid = bytes.len() == 38
        && bytes[0] == b'{'
        && bytes[37] == b'}'
        && [9, 14, 19, 24].iter().all(|index| bytes[*index] == b'-')
        && bytes.iter().enumerate().all(|(index, byte)| {
            matches!(index, 0 | 9 | 14 | 19 | 24 | 37) || byte.is_ascii_hexdigit()
        });
    if valid {
        Ok(())
    } else {
        Err(invalid("invalid section GUID"))
    }
}

fn children(node: &Node) -> Result<Vec<&Node>> {
    let mut children = Vec::new();
    for content in &node.content {
        match content {
            Content::Node(node) => children.push(node),
            Content::Text(text) if text.trim().is_empty() => {},
            Content::Comment(_) => {},
            _ => return Err(invalid("unexpected text in typed presentation sections")),
        }
    }
    Ok(children)
}

fn leaf(node: &Node) -> Result<()> {
    if children(node)?.is_empty() {
        Ok(())
    } else {
        Err(invalid("section leaf element has children"))
    }
}

fn expect(node: &Node, namespaces: &[&str], local: &str) -> Result<()> {
    if namespaces.contains(&node.namespace.as_str()) && node.local == local {
        Ok(())
    } else {
        Err(invalid(format!("expected {local}")))
    }
}

fn is(node: &Node, namespaces: &[&str], local: &str) -> bool {
    namespaces.contains(&node.namespace.as_str()) && node.local == local
}

fn optional_attr(node: &Node, local: &str) -> Result<Option<String>> {
    let mut value = None;
    for attribute in &node.attributes {
        if attribute.namespace.is_empty()
            && attribute.local == local
            && value.replace(attribute.value.clone()).is_some()
        {
            return Err(invalid(format!("duplicate attribute '{local}'")));
        }
    }
    Ok(value)
}

fn required_attr(node: &Node, local: &str) -> Result<String> {
    optional_attr(node, local)?.ok_or_else(|| invalid(format!("missing attribute '{local}'")))
}

fn only_unqualified(node: &Node, allowed: &[&str]) -> Result<()> {
    for attribute in &node.attributes {
        if !attribute.namespace.is_empty() || !allowed.contains(&attribute.local.as_str()) {
            return Err(invalid(format!(
                "unexpected attribute '{}'",
                attribute.qname
            )));
        }
    }
    Ok(())
}

fn no_attributes(node: &Node) -> Result<()> {
    only_unqualified(node, &[])
}

fn bounded_string(value: &str, limits: Limits) -> Result<()> {
    if value.len() > limits.max_string_bytes() {
        Err(limit(
            crate::opened::CodecLimitKind::StringBytes,
            value.len(),
            limits.max_string_bytes(),
        ))
    } else {
        Ok(())
    }
}

fn node_xml(node: &Node, limits: Limits) -> Result<Vec<u8>> {
    let mut xml =
        BoundedString::with_limit(limits.max_string_bytes(), "presentation section node XML");
    write_node(&mut xml, node)?;
    Ok(xml.finish().into_bytes())
}

fn write_node(xml: &mut BoundedString, node: &Node) -> Result<()> {
    xml.push('<')?;
    xml.push_str(&node.qname)?;
    for (prefix, uri) in &node.bindings {
        if prefix.is_empty() {
            xml.push_str(" xmlns=\"")?;
        } else {
            xml.push_str(" xmlns:")?;
            xml.push_str(prefix)?;
            xml.push_str("=\"")?;
        }
        escape_attribute(xml, uri)?;
        xml.push('"')?;
    }
    for attribute in &node.attributes {
        xml.push(' ')?;
        xml.push_str(&attribute.qname)?;
        xml.push_str("=\"")?;
        escape_attribute(xml, &attribute.value)?;
        xml.push('"')?;
    }
    if node.content.is_empty() {
        xml.push_str("/>")?;
        return Ok(());
    }
    xml.push('>')?;
    for content in &node.content {
        match content {
            Content::Node(node) => write_node(xml, node)?,
            Content::Text(text) => escape_text(xml, text)?,
            Content::CData(text) => {
                xml.push_str("<![CDATA[")?;
                xml.push_str(text)?;
                xml.push_str("]]>")?;
            },
            Content::Comment(text) => {
                xml.push_str("<!--")?;
                xml.push_str(text)?;
                xml.push_str("-->")?;
            },
        }
    }
    xml.push_str("</")?;
    xml.push_str(&node.qname)?;
    xml.push('>')?;
    Ok(())
}

fn split_qname(value: &str) -> Result<(&str, &str)> {
    if let Some((prefix, local)) = value.split_once(':') {
        if local.is_empty() || local.contains(':') {
            return Err(invalid("invalid QName"));
        }
        Ok((prefix, local))
    } else {
        Ok(("", value))
    }
}

fn resolve(bindings: &[(String, String)], prefix: &str) -> Result<String> {
    bindings
        .iter()
        .rev()
        .find(|binding| binding.0 == prefix)
        .map(|binding| binding.1.clone())
        .ok_or_else(|| invalid(format!("unbound namespace prefix '{prefix}'")))
}

fn escape_attribute(output: &mut BoundedString, value: &str) -> Result<()> {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;")?,
            '<' => output.push_str("&lt;")?,
            '"' => output.push_str("&quot;")?,
            '\r' => output.push_str("&#xD;")?,
            '\n' => output.push_str("&#xA;")?,
            '\t' => output.push_str("&#x9;")?,
            _ => output.push(character)?,
        }
    }
    Ok(())
}

fn escape_text(output: &mut BoundedString, value: &str) -> Result<()> {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;")?,
            '<' => output.push_str("&lt;")?,
            '>' => output.push_str("&gt;")?,
            _ => output.push(character)?,
        }
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(kind: crate::opened::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: crate::opened::LimitKind::Codec(kind),
        observed,
        maximum,
    }
}

fn check_empty_depth(depth: usize, limits: Limits) -> Result<()> {
    let observed = depth.checked_add(1).unwrap_or(usize::MAX);
    if observed > limits.max_depth() {
        return Err(limit(
            crate::opened::CodecLimitKind::XmlDepth,
            observed,
            limits.max_depth(),
        ));
    }
    Ok(())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    fn wrap(fragment: &str) -> String {
        format!(
            r#"<p:presentation xmlns:p="{P}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{fragment}</p:presentation>"#
        )
    }

    #[test]
    fn creates_and_round_trips_typed_sections() {
        let mut list = List::new();
        list.add_section(
            Section::new("Part 1", "{11111111-1111-1111-1111-111111111111}").with_slides([256]),
        );
        list.add_section(
            Section::new("Part 2", "{22222222-2222-2222-2222-222222222222}")
                .with_slides([257, 258]),
        );
        let xml = list.to_xml(Limits::DEFAULT).unwrap();
        assert_eq!(
            List::from_xml(wrap(&xml).as_bytes(), Limits::DEFAULT).unwrap(),
            list
        );
    }

    #[test]
    fn section_container_core_namespaces_match_the_presentation() {
        for (owner, other) in [(P, PS), (PS, P)] {
            let valid = format!(
                r#"<p:presentation xmlns:p="{owner}"><p:extLst><p:ext uri="urn:unrelated"/></p:extLst></p:presentation>"#
            );
            assert!(
                List::from_xml(valid.as_bytes(), Limits::DEFAULT)
                    .unwrap()
                    .sections()
                    .is_empty()
            );
            let mixed_list = format!(
                r#"<p:presentation xmlns:p="{owner}" xmlns:q="{other}"><q:extLst/></p:presentation>"#
            );
            assert!(List::from_xml(mixed_list.as_bytes(), Limits::DEFAULT).is_err());
            let mixed_entry = format!(
                r#"<p:presentation xmlns:p="{owner}" xmlns:q="{other}"><p:extLst><q:ext uri="urn:unrelated"/></p:extLst></p:presentation>"#
            );
            assert!(List::from_xml(mixed_entry.as_bytes(), Limits::DEFAULT).is_err());
        }
    }

    #[test]
    fn preserves_section_extensions_inertly_without_following_relationships() {
        let xml = wrap(&format!(
            r#"<p:extLst><p:ext uri="{SECTION_URI}"><p14:sectionLst xmlns:p14="{P14}"><p14:section name="Opaque" id="{{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}}"><p14:sldIdLst><p14:sldId id="256"/></p14:sldIdLst><p:extLst><p:ext uri="urn:producer"><v:payload xmlns:v="urn:vendor" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdNeverFetched" href="https://example.invalid/not-opened"/></p:ext></p:extLst></p14:section></p14:sectionLst></p:ext></p:extLst>"#
        ));
        let parsed = List::from_xml(xml.as_bytes(), Limits::DEFAULT).unwrap();
        let opaque =
            std::str::from_utf8(parsed.sections()[0].extension_xml.as_deref().unwrap()).unwrap();
        assert!(opaque.contains("rIdNeverFetched"));
        assert!(opaque.contains("https://example.invalid/not-opened"));
        let written = parsed.to_xml(Limits::DEFAULT).unwrap();
        let again = List::from_xml(wrap(&written).as_bytes(), Limits::DEFAULT).unwrap();
        let opaque =
            std::str::from_utf8(again.sections()[0].extension_xml.as_deref().unwrap()).unwrap();
        assert!(opaque.contains("rIdNeverFetched"));
        assert!(opaque.contains("https://example.invalid/not-opened"));
    }

    #[test]
    fn rejects_hostile_section_grammar_and_resource_cases() {
        let known = |body: &str| {
            wrap(&format!(
                r#"<p:extLst><p:ext uri="{SECTION_URI}">{body}</p:ext></p:extLst>"#
            ))
        };
        let cases = [
            known(&format!(r#"<p14:sectionLst xmlns:p14="{P14}"/>"#)),
            known(&format!(
                r#"<p14:sectionLst xmlns:p14="{P14}"><p14:section/></p14:sectionLst>"#
            )),
            known(&format!(
                r#"<p14:sectionLst xmlns:p14="{P14}"><p14:section id="not-a-guid"><p14:sldIdLst/></p14:section></p14:sectionLst>"#
            )),
            known(&format!(
                r#"<p14:sectionLst xmlns:p14="{P14}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p14:section r:id="rIdNo"><p14:sldIdLst/></p14:section></p14:sectionLst>"#
            )),
            known(&format!(
                r#"<p14:sectionLst xmlns:p14="{P14}"><p14:section><p14:sldIdLst><p14:sldId id="255"/></p14:sldIdLst></p14:section></p14:sectionLst>"#
            )),
            known(&format!(
                r#"<p14:sectionLst xmlns:p14="{P14}"><p14:section><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="256"/></p14:sldIdLst></p14:section></p14:sectionLst>"#
            )),
            known(&format!(
                r#"<p14:sectionLst xmlns:p14="{P14}" xmlns:v="urn:vendor"><v:section><p14:sldIdLst/></v:section></p14:sectionLst>"#
            )),
            format!(r#"<!DOCTYPE x><p:presentation xmlns:p="{P}"/>"#),
        ];
        for xml in cases {
            assert!(
                List::from_xml(xml.as_bytes(), Limits::DEFAULT).is_err(),
                "accepted {xml}"
            );
        }
        let duplicate = wrap(&format!(
            r#"<p:extLst><p:ext uri="{SECTION_URI}"><p14:sectionLst xmlns:p14="{P14}"><p14:section><p14:sldIdLst/></p14:section></p14:sectionLst></p:ext><p:ext uri="{SECTION_URI}"><p14:sectionLst xmlns:p14="{P14}"><p14:section><p14:sldIdLst/></p14:section></p14:sectionLst></p:ext></p:extLst>"#
        ));
        assert!(List::from_xml(duplicate.as_bytes(), Limits::DEFAULT).is_err());
        let mut oversized = List::new();
        for _ in 0..=Limits::DEFAULT.max_sections() {
            oversized.add_section(Section {
                name: None,
                id: None,
                slide_ids: Vec::new(),
                extension_xml: None,
            });
        }
        assert!(oversized.to_xml(Limits::DEFAULT).is_err());
    }

    #[test]
    fn empty_section_depth_is_checked_before_dom_ownership() {
        let limits = Limits::DEFAULT.with_depth(1).expect("depth limit");
        let exact = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;
        List::from_xml(exact, limits).expect("depth-one empty root");

        let one_over = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:holder/></p:presentation>"#;
        assert_xml_depth_limit(
            List::from_xml(one_over, limits).expect_err("depth-two empty element"),
            2,
            1,
        );
        assert!(matches!(
            List::from_xml(
                b"<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
                Limits::DEFAULT,
            ),
            Err(Error::Invalid(_))
        ));
    }

    fn assert_xml_depth_limit(error: Error, observed: usize, maximum: usize) {
        match error {
            Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlDepth),
                observed: actual,
                maximum: configured,
            } => {
                assert_eq!(actual, observed);
                assert_eq!(configured, maximum);
            },
            other => panic!("expected XML depth limit, got {other:?}"),
        }
    }
}
