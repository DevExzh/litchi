//! Typed, bounded support for PowerPoint 2010 presentation sections.

use crate::error::{OoxmlError, Result};
use litchi_core::xml::escape_xml;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
use std::collections::HashSet;
use std::fmt::Write as FmtWrite;

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const SECTION_URI: &str = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
const MAX_BYTES: usize = 8 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_NODES: usize = 100_000;
const MAX_SECTIONS: usize = 4_096;
const MAX_SLIDE_REFERENCES: usize = 100_000;
const MAX_EXTENSIONS: usize = 1_024;
const MAX_STRING_BYTES: usize = 1024 * 1024;

/// A logical group of presentation slides.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Section {
    /// Optional display name from `CT_Section`.
    pub name: Option<String>,
    /// Optional section GUID from `CT_Section`.
    pub id: Option<String>,
    /// Presentation slide identifiers in source order.
    pub slide_ids: Vec<u32>,
    /// Optional, inert `p:extLst` permitted by `CT_Section`.
    pub extension_xml: Option<Vec<u8>>,
}

impl Section {
    /// Create a named section with a GUID.
    pub fn new(name: impl Into<String>, id: impl Into<String>) -> Self {
        Self {
            name: Some(name.into()),
            id: Some(id.into()),
            slide_ids: Vec::new(),
            extension_xml: None,
        }
    }

    /// Add a presentation slide identifier.
    pub fn add_slide(&mut self, slide_id: u32) {
        self.slide_ids.push(slide_id);
    }

    /// Add presentation slide identifiers.
    pub fn with_slides(mut self, slide_ids: impl IntoIterator<Item = u32>) -> Self {
        self.slide_ids.extend(slide_ids);
        self
    }

    /// Serialize this `p14:section` element.
    pub fn to_xml(&self) -> Result<String> {
        validate_section(self)?;
        let mut xml = String::with_capacity(256);
        xml.push_str("<p14:section");
        if let Some(name) = &self.name {
            write!(xml, " name=\"{}\"", escape_xml(name))
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        }
        if let Some(id) = &self.id {
            write!(xml, " id=\"{}\"", escape_xml(id))
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        }
        xml.push_str("><p14:sldIdLst>");
        for slide_id in &self.slide_ids {
            write!(xml, "<p14:sldId id=\"{slide_id}\"/>")
                .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        }
        xml.push_str("</p14:sldIdLst>");
        if let Some(extension) = &self.extension_xml {
            xml.push_str(std::str::from_utf8(extension).map_err(xml_error)?);
        }
        xml.push_str("</p14:section>");
        Ok(xml)
    }
}

/// Ordered presentation sections.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SectionList {
    sections: Vec<Section>,
}

impl SectionList {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn add_section(&mut self, section: Section) {
        self.sections.push(section);
    }

    pub fn sections(&self) -> &[Section] {
        &self.sections
    }

    /// Get mutable access to the ordered sections.
    pub fn sections_mut(&mut self) -> &mut [Section] {
        &mut self.sections
    }

    /// Find a section by its stable GUID.
    pub fn get_by_id(&self, id: &str) -> Option<&Section> {
        self.sections
            .iter()
            .find(|section| section.id.as_deref() == Some(id))
    }

    /// Find a mutable section by its stable GUID.
    pub fn get_by_id_mut(&mut self, id: &str) -> Option<&mut Section> {
        self.sections
            .iter_mut()
            .find(|section| section.id.as_deref() == Some(id))
    }

    /// Replace a section while retaining its stable GUID.
    pub fn replace_by_id(&mut self, id: &str, mut replacement: Section) -> Result<()> {
        let target = self
            .get_by_id_mut(id)
            .ok_or_else(|| invalid(format!("section {id} was not found")))?;
        replacement.id = Some(id.to_owned());
        *target = replacement;
        validate_list(self)
    }

    /// Remove a section by its stable GUID.
    pub fn remove_by_id(&mut self, id: &str) -> Option<Section> {
        self.sections
            .iter()
            .position(|section| section.id.as_deref() == Some(id))
            .map(|offset| self.sections.remove(offset))
    }

    /// Reorder sections by a complete stable-GUID permutation.
    pub fn reorder(&mut self, ordered_ids: &[String]) -> Result<()> {
        let expected = self
            .sections
            .iter()
            .filter_map(|section| section.id.clone())
            .collect::<HashSet<_>>();
        let actual = ordered_ids.iter().cloned().collect::<HashSet<_>>();
        if expected.len() != self.sections.len()
            || expected != actual
            || ordered_ids.len() != self.sections.len()
        {
            return Err(invalid("section reorder is not a GUID permutation"));
        }
        self.sections = ordered_ids
            .iter()
            .map(|id| self.get_by_id(id).expect("permutation was validated").clone())
            .collect();
        Ok(())
    }

    /// Remove a slide from every section.
    pub(crate) fn remove_slide_membership(&mut self, slide_id: u32) {
        for section in &mut self.sections {
            section.slide_ids.retain(|id| *id != slide_id);
        }
    }

    /// Keep section membership in presentation order.
    pub(crate) fn sort_slide_membership(&mut self, ordered_slide_ids: &[u32]) {
        let positions = ordered_slide_ids
            .iter()
            .enumerate()
            .map(|(offset, id)| (*id, offset))
            .collect::<std::collections::HashMap<_, _>>();
        for section in &mut self.sections {
            section
                .slide_ids
                .sort_by_key(|id| positions.get(id).copied().unwrap_or(usize::MAX));
        }
    }

    pub fn len(&self) -> usize {
        self.sections.len()
    }

    pub fn is_empty(&self) -> bool {
        self.sections.is_empty()
    }

    /// Parse the section extension from a complete `p:presentation` document.
    ///
    /// Unrelated presentation extensions remain inert and are never evaluated.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_BYTES {
            return Err(invalid("presentation sections exceed 8 MiB"));
        }
        let processed = crate::common::mce::process_ooxml(xml)?;
        if processed.len() > MAX_BYTES {
            return Err(invalid("processed presentation sections exceed 8 MiB"));
        }
        parse_sections(&parse_dom(processed.as_ref())?)
    }

    /// Serialize the complete `p:extLst` fragment containing `p14:sectionLst`.
    pub fn to_xml(&self) -> Result<String> {
        if self.sections.is_empty() {
            return Ok(String::new());
        }
        validate_list(self)?;
        let mut xml = String::with_capacity(1024);
        xml.push_str("<p:extLst><p:ext uri=\"");
        xml.push_str(SECTION_URI);
        xml.push_str("\"><p14:sectionLst xmlns:p14=\"");
        xml.push_str(P14);
        xml.push_str("\">");
        for section in &self.sections {
            xml.push_str(&section.to_xml()?);
        }
        xml.push_str("</p14:sectionLst></p:ext></p:extLst>");
        if xml.len() > MAX_BYTES {
            return Err(invalid("serialized presentation sections exceed 8 MiB"));
        }
        Ok(xml)
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

fn parse_dom(xml: &[u8]) -> Result<Node> {
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut reader = Reader::from_reader(xml);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    loop {
        let decoder = reader.decoder();
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(invalid("presentation-section XML resource limit exceeded"));
                }
                stack.push(make_node(&element, decoder, &stack)?);
            },
            Ok(Event::Empty(element)) => {
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(invalid("presentation-section node limit exceeded"));
                }
                let node = make_node(&element, decoder, &stack)?;
                attach(&mut stack, &mut root, node)?;
            },
            Ok(Event::End(_)) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element"))?;
                attach(&mut stack, &mut root, node)?;
            },
            Ok(Event::Text(text)) => {
                let text = text.decode().map_err(xml_error)?.into_owned();
                if let Some(node) = stack.last_mut() {
                    node.content.push(Content::Text(text));
                } else if !text.trim().is_empty() {
                    return Err(invalid("text outside presentation root"));
                }
            },
            Ok(Event::CData(text)) => {
                let text = text.decode().map_err(xml_error)?.into_owned();
                if let Some(node) = stack.last_mut() {
                    node.content.push(Content::CData(text));
                } else {
                    return Err(invalid("CDATA outside presentation root"));
                }
            },
            Ok(Event::Comment(text)) => {
                if let Some(node) = stack.last_mut() {
                    node.content.push(Content::Comment(
                        text.decode().map_err(xml_error)?.into_owned(),
                    ));
                }
            },
            Ok(Event::GeneralRef(reference)) => {
                let text = crate::common::xml::decode_xml_reference(&reference)?;
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

fn make_node(element: &BytesStart<'_>, decoder: Decoder, stack: &[Node]) -> Result<Node> {
    let qname = std::str::from_utf8(element.name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    let mut bindings = stack
        .last()
        .map(|node| node.bindings.clone())
        .unwrap_or_default();
    let mut raw = Vec::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        raw.push((
            std::str::from_utf8(attribute.key.as_ref())
                .map_err(xml_error)?
                .to_owned(),
            attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(xml_error)?
                .into_owned(),
        ));
    }
    for (name, value) in &raw {
        if name == "xmlns" || name.starts_with("xmlns:") {
            let prefix = name.strip_prefix("xmlns:").unwrap_or("").to_owned();
            if let Some(binding) = bindings.iter_mut().find(|binding| binding.0 == prefix) {
                binding.1 = value.clone();
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

fn parse_sections(root: &Node) -> Result<SectionList> {
    expect(root, &[P, PS], "presentation")?;
    let mut presentation_ext = None;
    for child in children(root)? {
        if is(child, &[P, PS], "extLst") {
            if presentation_ext.replace(child).is_some() {
                return Err(invalid("duplicate presentation extLst"));
            }
        }
    }
    let Some(extension_list) = presentation_ext else {
        return Ok(SectionList::new());
    };
    let mut section_list = None;
    for extension in children(extension_list)? {
        expect(extension, &[P, PS], "ext")?;
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
            section_list = Some(parse_section_list(payload[0])?);
        }
    }
    Ok(section_list.unwrap_or_default())
}

fn parse_section_list(node: &Node) -> Result<SectionList> {
    expect(node, &[P14], "sectionLst")?;
    no_attributes(node)?;
    let sections = children(node)?;
    if sections.is_empty() {
        return Err(invalid("sectionLst requires at least one section"));
    }
    if sections.len() > MAX_SECTIONS {
        return Err(invalid("presentation section count exceeds limit"));
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
            bounded_string(name)?;
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
        let slide_ids = parse_slide_id_list(content[0], &mut total_slide_ids)?;
        let extension_xml = if content.len() == 2 {
            validate_extension_list(content[1])?;
            Some(node_xml(content[1])?)
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
    Ok(SectionList { sections: parsed })
}

fn parse_slide_id_list(node: &Node, total: &mut usize) -> Result<Vec<u32>> {
    no_attributes(node)?;
    let mut slide_ids = Vec::new();
    let mut seen = HashSet::new();
    for slide in children(node)? {
        expect(slide, &[P14], "sldId")?;
        let id = required_attr(slide, "id")?
            .parse::<u32>()
            .map_err(|_| invalid("invalid section slide ID"))?;
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
        if *total > MAX_SLIDE_REFERENCES {
            return Err(invalid("section slide reference count exceeds limit"));
        }
        slide_ids.push(id);
    }
    Ok(slide_ids)
}

fn validate_extension_list(node: &Node) -> Result<()> {
    expect(node, &[P, PS], "extLst")?;
    no_attributes(node)?;
    let extensions = children(node)?;
    if extensions.len() > MAX_EXTENSIONS {
        return Err(invalid("section extension count exceeds limit"));
    }
    for extension in extensions {
        expect(extension, &[P, PS], "ext")?;
        let uri = required_attr(extension, "uri")?;
        if uri.is_empty() {
            return Err(invalid("section extension URI is empty"));
        }
        bounded_string(&uri)?;
        only_unqualified(extension, &["uri"])?;
    }
    Ok(())
}

fn validate_list(list: &SectionList) -> Result<()> {
    if list.sections.len() > MAX_SECTIONS {
        return Err(invalid("presentation section count exceeds limit"));
    }
    let mut section_ids = HashSet::new();
    let mut total = 0usize;
    for section in &list.sections {
        validate_section(section)?;
        if let Some(id) = &section.id {
            if !section_ids.insert(id) {
                return Err(invalid("duplicate section GUID"));
            }
        }
        total = total
            .checked_add(section.slide_ids.len())
            .ok_or_else(|| invalid("section slide reference count overflow"))?;
        if total > MAX_SLIDE_REFERENCES {
            return Err(invalid("section slide reference count exceeds limit"));
        }
    }
    Ok(())
}

fn validate_section(section: &Section) -> Result<()> {
    if let Some(name) = &section.name {
        bounded_string(name)?;
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
        if extension.len() > MAX_BYTES {
            return Err(invalid("section extension XML exceeds limit"));
        }
        validate_extension_list(&parse_dom(extension)?)?;
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
        if attribute.namespace.is_empty() && attribute.local == local {
            if value.replace(attribute.value.clone()).is_some() {
                return Err(invalid(format!("duplicate attribute '{local}'")));
            }
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

fn bounded_string(value: &str) -> Result<()> {
    if value.len() > MAX_STRING_BYTES {
        Err(invalid("presentation section string exceeds 1 MiB"))
    } else {
        Ok(())
    }
}

fn node_xml(node: &Node) -> Result<Vec<u8>> {
    let mut xml = String::new();
    write_node(&mut xml, node)?;
    Ok(xml.into_bytes())
}

fn write_node(xml: &mut String, node: &Node) -> Result<()> {
    xml.push('<');
    xml.push_str(&node.qname);
    for (prefix, uri) in &node.bindings {
        if prefix.is_empty() {
            xml.push_str(" xmlns=\"");
        } else {
            xml.push_str(" xmlns:");
            xml.push_str(prefix);
            xml.push_str("=\"");
        }
        escape_attribute(xml, uri);
        xml.push('"');
    }
    for attribute in &node.attributes {
        xml.push(' ');
        xml.push_str(&attribute.qname);
        xml.push_str("=\"");
        escape_attribute(xml, &attribute.value);
        xml.push('"');
    }
    if node.content.is_empty() {
        xml.push_str("/>");
        return Ok(());
    }
    xml.push('>');
    for content in &node.content {
        match content {
            Content::Node(node) => write_node(xml, node)?,
            Content::Text(text) => escape_text(xml, text),
            Content::CData(text) => {
                xml.push_str("<![CDATA[");
                xml.push_str(text);
                xml.push_str("]]>");
            },
            Content::Comment(text) => {
                xml.push_str("<!--");
                xml.push_str(text);
                xml.push_str("-->");
            },
        }
    }
    xml.push_str("</");
    xml.push_str(&node.qname);
    xml.push('>');
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

fn escape_attribute(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '"' => output.push_str("&quot;"),
            '\r' => output.push_str("&#xD;"),
            '\n' => output.push_str("&#xA;"),
            '\t' => output.push_str("&#x9;"),
            _ => output.push(character),
        }
    }
}

fn escape_text(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            _ => output.push(character),
        }
    }
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pptx::parts::PresentationPart;
    use litchi_opc::OpcPackage;
    use litchi_opc::packuri::PackURI;
    use litchi_opc::part::BlobPart;

    const LO_SECTIONS: &[u8] = include_bytes!(
        "../../../../test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-sections.pptx"
    );
    const LO_SECTION_TEST: &[u8] = include_bytes!(
        "../../../../test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx"
    );

    fn wrap(fragment: &str) -> String {
        format!(
            r#"<p:presentation xmlns:p="{P}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{fragment}</p:presentation>"#
        )
    }

    fn part(xml: impl Into<Vec<u8>>) -> BlobPart {
        BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                .to_owned(),
            xml.into(),
        )
    }

    #[test]
    fn creates_and_round_trips_typed_sections() {
        let mut list = SectionList::new();
        list.add_section(
            Section::new("Part 1", "{11111111-1111-1111-1111-111111111111}").with_slides([256]),
        );
        list.add_section(
            Section::new("Part 2", "{22222222-2222-2222-2222-222222222222}")
                .with_slides([257, 258]),
        );
        let xml = list.to_xml().unwrap();
        assert_eq!(SectionList::from_xml(wrap(&xml).as_bytes()).unwrap(), list);
    }

    #[test]
    fn loads_libreoffice_section_packages_and_validates_membership() {
        for (bytes, count, first_name, first_len) in [
            (LO_SECTIONS, 2, "Default Section", 3),
            (LO_SECTION_TEST, 3, "Section-1", 4),
        ] {
            let package = OpcPackage::from_bytes(bytes).unwrap();
            let presentation =
                PresentationPart::from_part(package.main_document_part().unwrap()).unwrap();
            let sections = presentation.sections().unwrap();
            assert_eq!(sections.len(), count);
            assert_eq!(sections.sections()[0].name.as_deref(), Some(first_name));
            assert_eq!(sections.sections()[0].slide_ids.len(), first_len);
            let fragment = sections.to_xml().unwrap();
            assert_eq!(
                SectionList::from_xml(wrap(&fragment).as_bytes()).unwrap(),
                sections
            );
        }
    }

    #[test]
    fn preserves_section_extensions_inertly_without_following_relationships() {
        let xml = wrap(&format!(
            r#"<p:extLst><p:ext uri="{SECTION_URI}"><p14:sectionLst xmlns:p14="{P14}"><p14:section name="Opaque" id="{{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}}"><p14:sldIdLst><p14:sldId id="256"/></p14:sldIdLst><p:extLst><p:ext uri="urn:producer"><v:payload xmlns:v="urn:vendor" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdNeverFetched" href="https://example.invalid/not-opened"/></p:ext></p:extLst></p14:section></p14:sectionLst></p:ext></p:extLst>"#
        ));
        let parsed = SectionList::from_xml(xml.as_bytes()).unwrap();
        let opaque =
            std::str::from_utf8(parsed.sections()[0].extension_xml.as_deref().unwrap()).unwrap();
        assert!(opaque.contains("rIdNeverFetched"));
        assert!(opaque.contains("https://example.invalid/not-opened"));
        let written = parsed.to_xml().unwrap();
        let again = SectionList::from_xml(wrap(&written).as_bytes()).unwrap();
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
                SectionList::from_xml(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
        let duplicate = wrap(&format!(
            r#"<p:extLst><p:ext uri="{SECTION_URI}"><p14:sectionLst xmlns:p14="{P14}"><p14:section><p14:sldIdLst/></p14:section></p14:sectionLst></p:ext><p:ext uri="{SECTION_URI}"><p14:sectionLst xmlns:p14="{P14}"><p14:section><p14:sldIdLst/></p14:section></p14:sectionLst></p:ext></p:extLst>"#
        ));
        assert!(SectionList::from_xml(duplicate.as_bytes()).is_err());
        let mut oversized = SectionList::new();
        for _ in 0..=MAX_SECTIONS {
            oversized.add_section(Section {
                name: None,
                id: None,
                slide_ids: Vec::new(),
                extension_xml: None,
            });
        }
        assert!(oversized.to_xml().is_err());
    }

    #[test]
    fn presentation_part_rejects_section_references_to_undeclared_slides() {
        let xml = format!(
            r#"<p:presentation xmlns:p="{P}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:extLst><p:ext uri="{SECTION_URI}"><p14:sectionLst xmlns:p14="{P14}"><p14:section><p14:sldIdLst><p14:sldId id="257"/></p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst></p:presentation>"#
        );
        assert!(
            PresentationPart::from_part(&part(xml))
                .unwrap()
                .sections()
                .is_err()
        );
    }
}
