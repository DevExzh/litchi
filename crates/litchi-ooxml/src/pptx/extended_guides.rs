//! Typed, bounded PowerPoint 2013 extended presentation guides.

use crate::error::{OoxmlError, Result};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
use std::collections::HashSet;
use std::fmt::Write;

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const AS: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const P15: &str = "http://schemas.microsoft.com/office/powerpoint/2012/main";
const SLIDE_GUIDES_URI: &str = "{EFAFB233-063F-42B5-8137-9DF3F51BA10A}";
const NOTES_GUIDES_URI: &str = "{2D200454-40CA-4A62-9FC3-DE9A4176ACB9}";
const MAX_BYTES: usize = 8 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_NODES: usize = 100_000;
const MAX_GUIDES: usize = 16_384;
const MAX_EXTENSIONS: usize = 1_024;
const MAX_STRING_BYTES: usize = 1024 * 1024;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ExtendedGuideOrientation {
    Horizontal,
    Vertical,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ExtendedGuideColorKind {
    ScRgb,
    Srgb,
    Hsl,
    System,
    Scheme,
    Preset,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ExtendedGuideColor {
    pub kind: ExtendedGuideColorKind,
    /// Inert DrawingML color XML, including transforms.
    pub xml: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ExtendedGuide {
    pub id: u32,
    pub name: Option<String>,
    pub orientation: Option<ExtendedGuideOrientation>,
    pub position: Option<i32>,
    pub user_drawn: Option<bool>,
    pub color: ExtendedGuideColor,
    /// Optional, inert `p:extLst` permitted by `CT_ExtendedGuide`.
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ExtendedGuideList {
    pub guides: Vec<ExtendedGuide>,
    /// Optional, inert `p:extLst` permitted by `CT_ExtendedGuideList`.
    pub extension_xml: Option<Vec<u8>>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct PresentationExtendedGuides {
    pub slide: Option<ExtendedGuideList>,
    pub notes: Option<ExtendedGuideList>,
}

impl PresentationExtendedGuides {
    /// Parse guide extensions from a complete `p:presentation` document.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_BYTES {
            return Err(invalid("presentation guides exceed 8 MiB"));
        }
        let processed = crate::common::mce::process_ooxml(xml)?;
        if processed.len() > MAX_BYTES {
            return Err(invalid("processed presentation guides exceed 8 MiB"));
        }
        parse_presentation(&parse_dom(processed.as_ref())?)
    }

    /// Serialize the guide entries as a complete `p:extLst` fragment.
    pub fn to_xml(&self, strict: bool) -> Result<String> {
        validate(self)?;
        if self.slide.is_none() && self.notes.is_none() {
            return Ok(String::new());
        }
        let mut xml = String::with_capacity(1024);
        xml.push_str("<p:extLst>");
        if let Some(list) = &self.slide {
            write_list_extension(&mut xml, SLIDE_GUIDES_URI, "sldGuideLst", list, strict)?;
        }
        if let Some(list) = &self.notes {
            write_list_extension(&mut xml, NOTES_GUIDES_URI, "notesGuideLst", list, strict)?;
        }
        xml.push_str("</p:extLst>");
        if xml.len() > MAX_BYTES {
            return Err(invalid("serialized presentation guides exceed 8 MiB"));
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
    let mut count = 0usize;
    loop {
        let decoder = reader.decoder();
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                count += 1;
                if count > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(invalid("presentation-guide XML resource limit exceeded"));
                }
                stack.push(make_node(&element, decoder, &stack)?);
            },
            Ok(Event::Empty(element)) => {
                count += 1;
                if count > MAX_NODES {
                    return Err(invalid("presentation-guide node limit exceeded"));
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
        return Err(invalid("unterminated presentation-guide XML"));
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

fn parse_presentation(root: &Node) -> Result<PresentationExtendedGuides> {
    expect(root, &[P, PS], "presentation")?;
    let mut root_ext = None;
    for child in children(root)? {
        if is(child, &[P, PS], "extLst") && root_ext.replace(child).is_some() {
            return Err(invalid("duplicate presentation extLst"));
        }
    }
    let Some(root_ext) = root_ext else {
        return Ok(PresentationExtendedGuides::default());
    };
    let mut value = PresentationExtendedGuides::default();
    for extension in children(root_ext)? {
        expect(extension, &[P, PS], "ext")?;
        let uri = required_attr(extension, "uri")?;
        let target = match uri.as_str() {
            SLIDE_GUIDES_URI => Some((&mut value.slide, "sldGuideLst")),
            NOTES_GUIDES_URI => Some((&mut value.notes, "notesGuideLst")),
            _ => None,
        };
        if let Some((slot, local)) = target {
            if slot.is_some() {
                return Err(invalid(format!("duplicate {local} extension")));
            }
            only_unqualified(extension, &["uri"])?;
            let payload = children(extension)?;
            if payload.len() != 1 {
                return Err(invalid(format!("{local} extension requires one payload")));
            }
            expect(payload[0], &[P15], local)?;
            *slot = Some(parse_list(payload[0])?);
        }
    }
    Ok(value)
}

fn parse_list(node: &Node) -> Result<ExtendedGuideList> {
    no_attributes(node)?;
    let content = children(node)?;
    let mut guides = Vec::new();
    let mut extension_xml = None;
    let mut ids = HashSet::new();
    for child in content {
        if is(child, &[P15], "guide") {
            if extension_xml.is_some() {
                return Err(invalid("guide appears after guide-list extLst"));
            }
            if guides.len() >= MAX_GUIDES {
                return Err(invalid("extended guide count exceeds limit"));
            }
            let guide = parse_guide(child)?;
            if !ids.insert(guide.id) {
                return Err(invalid("duplicate extended guide ID"));
            }
            guides.push(guide);
        } else if is(child, &[P, PS], "extLst") {
            if extension_xml.is_some() {
                return Err(invalid("duplicate guide-list extLst"));
            }
            validate_extension_list(child)?;
            extension_xml = Some(node_xml(child, false)?);
        } else {
            return Err(invalid("unexpected extended guide-list child"));
        }
    }
    Ok(ExtendedGuideList {
        guides,
        extension_xml,
    })
}

fn parse_guide(node: &Node) -> Result<ExtendedGuide> {
    let id = required_attr(node, "id")?
        .parse::<u32>()
        .map_err(|_| invalid("invalid extended guide ID"))?;
    let name = optional_attr(node, "name")?;
    if let Some(name) = &name {
        bounded_string(name)?;
    }
    let orientation = optional_attr(node, "orient")?
        .map(|value| match value.as_str() {
            "horz" => Ok(ExtendedGuideOrientation::Horizontal),
            "vert" => Ok(ExtendedGuideOrientation::Vertical),
            _ => Err(invalid("invalid extended guide orientation")),
        })
        .transpose()?;
    let position = optional_attr(node, "pos")?
        .map(|value| {
            value
                .parse::<i32>()
                .map_err(|_| invalid("invalid extended guide position"))
        })
        .transpose()?;
    let user_drawn = optional_attr(node, "userDrawn")?
        .map(|value| parse_bool(&value, "userDrawn"))
        .transpose()?;
    only_unqualified(node, &["id", "name", "orient", "pos", "userDrawn"])?;
    let content = children(node)?;
    if content.is_empty()
        || content.len() > 2
        || !is(content[0], &[P15], "clr")
        || (content.len() == 2 && !is(content[1], &[P, PS], "extLst"))
    {
        return Err(invalid("guide requires clr followed by optional extLst"));
    }
    let color = parse_color(content[0])?;
    let extension_xml = if content.len() == 2 {
        validate_extension_list(content[1])?;
        Some(node_xml(content[1], false)?)
    } else {
        None
    };
    Ok(ExtendedGuide {
        id,
        name,
        orientation,
        position,
        user_drawn,
        color,
        extension_xml,
    })
}

fn parse_color(node: &Node) -> Result<ExtendedGuideColor> {
    no_attributes(node)?;
    let colors = children(node)?;
    if colors.len() != 1 {
        return Err(invalid("extended guide clr requires one DrawingML color"));
    }
    let color = colors[0];
    let kind = match (color.namespace.as_str(), color.local.as_str()) {
        (A | AS, "scrgbClr") => ExtendedGuideColorKind::ScRgb,
        (A | AS, "srgbClr") => ExtendedGuideColorKind::Srgb,
        (A | AS, "hslClr") => ExtendedGuideColorKind::Hsl,
        (A | AS, "sysClr") => ExtendedGuideColorKind::System,
        (A | AS, "schemeClr") => ExtendedGuideColorKind::Scheme,
        (A | AS, "prstClr") => ExtendedGuideColorKind::Preset,
        _ => return Err(invalid("invalid extended guide DrawingML color")),
    };
    Ok(ExtendedGuideColor {
        kind,
        xml: node_xml(color, false)?,
    })
}

fn validate_extension_list(node: &Node) -> Result<()> {
    expect(node, &[P, PS], "extLst")?;
    no_attributes(node)?;
    let extensions = children(node)?;
    if extensions.len() > MAX_EXTENSIONS {
        return Err(invalid("guide extension count exceeds limit"));
    }
    for extension in extensions {
        expect(extension, &[P, PS], "ext")?;
        let uri = required_attr(extension, "uri")?;
        if uri.is_empty() {
            return Err(invalid("guide extension URI is empty"));
        }
        bounded_string(&uri)?;
        only_unqualified(extension, &["uri"])?;
    }
    Ok(())
}

fn validate(value: &PresentationExtendedGuides) -> Result<()> {
    for list in value.slide.iter().chain(value.notes.iter()) {
        if list.guides.len() > MAX_GUIDES {
            return Err(invalid("extended guide count exceeds limit"));
        }
        let mut ids = HashSet::new();
        for guide in &list.guides {
            if !ids.insert(guide.id) {
                return Err(invalid("duplicate extended guide ID"));
            }
            if let Some(name) = &guide.name {
                bounded_string(name)?;
            }
            validate_color(&guide.color)?;
            if let Some(extension) = &guide.extension_xml {
                validate_opaque_extension(extension)?;
            }
        }
        if let Some(extension) = &list.extension_xml {
            validate_opaque_extension(extension)?;
        }
    }
    Ok(())
}

fn validate_color(color: &ExtendedGuideColor) -> Result<()> {
    if color.xml.len() > MAX_BYTES {
        return Err(invalid("extended guide color XML exceeds limit"));
    }
    let node = parse_dom(&color.xml)?;
    let parsed = parse_color_node(&node)?;
    if parsed != color.kind {
        return Err(invalid("extended guide color kind does not match XML"));
    }
    Ok(())
}

fn parse_color_node(node: &Node) -> Result<ExtendedGuideColorKind> {
    match (node.namespace.as_str(), node.local.as_str()) {
        (A | AS, "scrgbClr") => Ok(ExtendedGuideColorKind::ScRgb),
        (A | AS, "srgbClr") => Ok(ExtendedGuideColorKind::Srgb),
        (A | AS, "hslClr") => Ok(ExtendedGuideColorKind::Hsl),
        (A | AS, "sysClr") => Ok(ExtendedGuideColorKind::System),
        (A | AS, "schemeClr") => Ok(ExtendedGuideColorKind::Scheme),
        (A | AS, "prstClr") => Ok(ExtendedGuideColorKind::Preset),
        _ => Err(invalid("invalid extended guide DrawingML color")),
    }
}

fn validate_opaque_extension(xml: &[u8]) -> Result<()> {
    if xml.len() > MAX_BYTES {
        return Err(invalid("guide extension XML exceeds limit"));
    }
    validate_extension_list(&parse_dom(xml)?)
}

fn write_list_extension(
    xml: &mut String,
    uri: &str,
    local: &str,
    list: &ExtendedGuideList,
    strict: bool,
) -> Result<()> {
    write!(
        xml,
        "<p:ext uri=\"{uri}\"><p15:{local} xmlns:p15=\"{P15}\">"
    )
    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    for guide in &list.guides {
        write_guide(xml, guide, strict)?;
    }
    if let Some(extension) = &list.extension_xml {
        write_opaque(xml, extension, strict)?;
    }
    write!(xml, "</p15:{local}></p:ext>").map_err(|error| OoxmlError::Xml(error.to_string()))?;
    Ok(())
}

fn write_guide(xml: &mut String, guide: &ExtendedGuide, strict: bool) -> Result<()> {
    write!(xml, "<p15:guide id=\"{}\"", guide.id)
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    if let Some(name) = &guide.name {
        xml.push_str(" name=\"");
        escape_attribute(xml, name);
        xml.push('"');
    }
    if let Some(orientation) = guide.orientation {
        xml.push_str(" orient=\"");
        xml.push_str(match orientation {
            ExtendedGuideOrientation::Horizontal => "horz",
            ExtendedGuideOrientation::Vertical => "vert",
        });
        xml.push('"');
    }
    if let Some(position) = guide.position {
        write!(xml, " pos=\"{position}\"").map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    if let Some(user_drawn) = guide.user_drawn {
        write!(xml, " userDrawn=\"{}\"", if user_drawn { 1 } else { 0 })
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    }
    xml.push_str("><p15:clr>");
    write_opaque(xml, &guide.color.xml, strict)?;
    xml.push_str("</p15:clr>");
    if let Some(extension) = &guide.extension_xml {
        write_opaque(xml, extension, strict)?;
    }
    xml.push_str("</p15:guide>");
    Ok(())
}

fn write_opaque(output: &mut String, xml: &[u8], strict: bool) -> Result<()> {
    let mut text = std::str::from_utf8(xml).map_err(xml_error)?.to_owned();
    if strict {
        text = text.replace(P, PS).replace(A, AS);
    } else {
        text = text.replace(PS, P).replace(AS, A);
    }
    output.push_str(&text);
    Ok(())
}

fn children(node: &Node) -> Result<Vec<&Node>> {
    let mut children = Vec::new();
    for content in &node.content {
        match content {
            Content::Node(node) => children.push(node),
            Content::Text(text) if text.trim().is_empty() => {},
            Content::Comment(_) => {},
            _ => return Err(invalid("unexpected text in typed presentation guides")),
        }
    }
    Ok(children)
}

fn expect(node: &Node, namespaces: &[&str], local: &str) -> Result<()> {
    if is(node, namespaces, local) {
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
        Err(invalid("presentation guide string exceeds 1 MiB"))
    } else {
        Ok(())
    }
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid boolean '{name}'"))),
    }
}

fn node_xml(node: &Node, strict: bool) -> Result<Vec<u8>> {
    let mut xml = String::new();
    write_node(&mut xml, node, strict)?;
    Ok(xml.into_bytes())
}

fn write_node(xml: &mut String, node: &Node, strict: bool) -> Result<()> {
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
        let uri = if strict {
            match uri.as_str() {
                P => PS,
                A => AS,
                _ => uri,
            }
        } else {
            match uri.as_str() {
                PS => P,
                AS => A,
                _ => uri,
            }
        };
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
            Content::Node(node) => write_node(xml, node, strict)?,
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

    const LO_GUIDES: &[u8] = include_bytes!(
        "../../../../test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-sections.pptx"
    );

    fn wrap(fragment: &str) -> String {
        format!(
            r#"<p:presentation xmlns:p="{P}" xmlns:a="{A}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{fragment}</p:presentation>"#
        )
    }

    #[test]
    fn loads_libreoffice_extended_guides_from_package() {
        let package = OpcPackage::from_bytes(LO_GUIDES).unwrap();
        let presentation =
            PresentationPart::from_part(package.main_document_part().unwrap()).unwrap();
        let value = presentation.extended_guides().unwrap();
        let guides = &value.slide.as_ref().unwrap().guides;
        assert_eq!(guides.len(), 2);
        assert_eq!(guides[0].id, 1);
        assert_eq!(
            guides[0].orientation,
            Some(ExtendedGuideOrientation::Horizontal)
        );
        assert_eq!(guides[0].position, Some(2160));
        assert_eq!(guides[0].user_drawn, Some(true));
        assert_eq!(guides[0].color.kind, ExtendedGuideColorKind::Srgb);
        assert_eq!(guides[1].orientation, None);
        let xml = value.to_xml(false).unwrap();
        let again = PresentationExtendedGuides::from_xml(wrap(&xml).as_bytes()).unwrap();
        assert_eq!(again.slide.unwrap().guides.len(), 2);
    }

    #[test]
    fn round_trips_slide_notes_and_unknown_extensions_inertly() {
        let xml = wrap(&format!(
            r#"<p:extLst><p:ext uri="{SLIDE_GUIDES_URI}"><p15:sldGuideLst xmlns:p15="{P15}"><p15:guide id="7" name="Named" orient="vert" pos="-20" userDrawn="0"><p15:clr><a:schemeClr val="accent1"/></p15:clr><p:extLst><p:ext uri="urn:guide"><v:data xmlns:v="urn:vendor" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdNeverFetched" href="https://example.invalid/not-opened"/></p:ext></p:extLst></p15:guide><p:extLst><p:ext uri="urn:list"><v:list xmlns:v="urn:vendor"/></p:ext></p:extLst></p15:sldGuideLst></p:ext><p:ext uri="{NOTES_GUIDES_URI}"><p15:notesGuideLst xmlns:p15="{P15}"/></p:ext></p:extLst>"#
        ));
        let value = PresentationExtendedGuides::from_xml(xml.as_bytes()).unwrap();
        let guide = &value.slide.as_ref().unwrap().guides[0];
        assert_eq!(guide.name.as_deref(), Some("Named"));
        assert_eq!(guide.color.kind, ExtendedGuideColorKind::Scheme);
        let opaque = std::str::from_utf8(guide.extension_xml.as_deref().unwrap()).unwrap();
        assert!(opaque.contains("rIdNeverFetched"));
        assert!(opaque.contains("https://example.invalid/not-opened"));
        assert!(value.notes.as_ref().unwrap().guides.is_empty());
        for strict in [false, true] {
            let written = value.to_xml(strict).unwrap();
            let again = PresentationExtendedGuides::from_xml(wrap(&written).as_bytes()).unwrap();
            assert_eq!(again.slide.as_ref().unwrap().guides[0].id, 7);
            assert!(
                std::str::from_utf8(
                    again.slide.as_ref().unwrap().guides[0]
                        .extension_xml
                        .as_deref()
                        .unwrap()
                )
                .unwrap()
                .contains("rIdNeverFetched")
            );
        }
    }

    #[test]
    fn rejects_hostile_extended_guide_grammar() {
        let known = |body: &str| {
            wrap(&format!(
                r#"<p:extLst><p:ext uri="{SLIDE_GUIDES_URI}">{body}</p:ext></p:extLst>"#
            ))
        };
        let cases = [
            known(&format!(
                r#"<p15:sldGuideLst xmlns:p15="{P15}"><p15:guide><p15:clr><a:srgbClr val="AABBCC"/></p15:clr></p15:guide></p15:sldGuideLst>"#
            )),
            known(&format!(
                r#"<p15:sldGuideLst xmlns:p15="{P15}"><p15:guide id="1" orient="diagonal"><p15:clr><a:srgbClr val="AABBCC"/></p15:clr></p15:guide></p15:sldGuideLst>"#
            )),
            known(&format!(
                r#"<p15:sldGuideLst xmlns:p15="{P15}"><p15:guide id="1"><p15:clr/></p15:guide></p15:sldGuideLst>"#
            )),
            known(&format!(
                r#"<p15:sldGuideLst xmlns:p15="{P15}"><p15:guide id="1"><p15:clr><a:srgbClr val="AABBCC"/><a:schemeClr val="accent1"/></p15:clr></p15:guide></p15:sldGuideLst>"#
            )),
            known(&format!(
                r#"<p15:sldGuideLst xmlns:p15="{P15}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p15:guide id="1" r:id="rIdNo"><p15:clr><a:srgbClr val="AABBCC"/></p15:clr></p15:guide></p15:sldGuideLst>"#
            )),
            known(&format!(
                r#"<p15:sldGuideLst xmlns:p15="{P15}"><p15:guide id="1"><p15:clr><a:srgbClr val="AABBCC"/></p15:clr></p15:guide><p15:guide id="1"><p15:clr><a:srgbClr val="AABBCC"/></p15:clr></p15:guide></p15:sldGuideLst>"#
            )),
            known(&format!(r#"<p15:notesGuideLst xmlns:p15="{P15}"/>"#)),
            format!(r#"<!DOCTYPE x><p:presentation xmlns:p="{P}"/>"#),
        ];
        for xml in cases {
            assert!(
                PresentationExtendedGuides::from_xml(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
    }

    #[test]
    fn rejects_programmatic_limits_and_color_mismatch() {
        let color = ExtendedGuideColor {
            kind: ExtendedGuideColorKind::Srgb,
            xml: format!(r#"<a:srgbClr xmlns:a="{A}" val="AABBCC"/>"#).into_bytes(),
        };
        let guide = ExtendedGuide {
            id: 1,
            name: None,
            orientation: None,
            position: None,
            user_drawn: None,
            color: color.clone(),
            extension_xml: None,
        };
        let mut value = PresentationExtendedGuides {
            slide: Some(ExtendedGuideList {
                guides: vec![guide; MAX_GUIDES + 1],
                extension_xml: None,
            }),
            notes: None,
        };
        assert!(value.to_xml(false).is_err());
        value.slide.as_mut().unwrap().guides.truncate(1);
        value.slide.as_mut().unwrap().guides[0].color.kind = ExtendedGuideColorKind::Scheme;
        assert!(value.to_xml(false).is_err());
    }
}
