//! Typed, bounded PowerPoint 2013 extended presentation guides.

use super::model::*;
use crate::{Error, Result};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
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
pub(crate) const MAX_BYTES: usize = 8 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_NODES: usize = 100_000;
const MAX_GUIDES: usize = 16_384;
const MAX_EXTENSIONS: usize = 1_024;
const MAX_STRING_BYTES: usize = 1024 * 1024;

impl Guides {
    /// Parse guide extensions from a complete `p:presentation` document.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_BYTES {
            return Err(invalid("presentation guides exceed 8 MiB"));
        }
        let processed = litchi_ooxml_common::mce::process_ooxml(xml)?;
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

/// Validate a programmatically staged guide value before package publication.
pub(crate) fn validate_value(value: &Guides) -> Result<()> {
    validate(value)
}

/// Replace only the known guide extension entries in one presentation source.
///
/// The source is preprocessed through the shared MCE baseline before a changed
/// edit is written. The no-op transaction path deliberately skips this helper,
/// so an untouched source is returned byte-for-byte unchanged.
pub(crate) fn rewrite_source(source: &[u8], value: &Guides) -> Result<Vec<u8>> {
    if source.len() > MAX_BYTES {
        return Err(invalid("presentation guides exceed 8 MiB"));
    }
    validate(value)?;
    let processed = litchi_ooxml_common::mce::process_ooxml(source)?;
    if processed.len() > MAX_BYTES {
        return Err(invalid("processed presentation guides exceed 8 MiB"));
    }
    let source = processed.as_ref();
    let layout = scan_source(source)?;
    let entries = target_entries(value, layout.strict, layout.presentation_namespace)?;

    let Some(ext_list) = layout.ext_list.as_ref() else {
        if entries.iter().all(Option::is_none) {
            return Ok(source.to_vec());
        }
        let fragment = format!(
            "<p:extLst xmlns:p=\"{}\">{}</p:extLst>",
            layout.presentation_namespace,
            entries
                .iter()
                .filter_map(Option::as_deref)
                .collect::<String>()
        );
        return if layout.root.empty {
            let replacement = expand_empty(source, &layout.root, &fragment)?;
            replace_spans(source, &[(layout.root.start, layout.root.end, replacement)])
        } else {
            let root_close = layout
                .root_close
                .ok_or_else(|| invalid("presentation root is missing its end tag"))?;
            insert_bytes(source, root_close, fragment.as_bytes())
        };
    };

    let mut replacements = Vec::new();
    for (index, _) in [Target::Slide, Target::Notes].into_iter().enumerate() {
        let replacement = entries[index].as_deref().unwrap_or_default().as_bytes();
        if let Some(span) = layout.targets[index].as_ref() {
            replacements.push((span.start, span.end, replacement.to_vec()));
        } else if !replacement.is_empty() {
            if ext_list.empty {
                let body = entries
                    .iter()
                    .filter_map(|entry| entry.as_deref())
                    .collect::<String>();
                let replacement = expand_empty_bytes(source, ext_list, body.as_bytes())?;
                return replace_spans(source, &[(ext_list.start, ext_list.end, replacement)]);
            }
            replacements.push((
                ext_list.close_start,
                ext_list.close_start,
                replacement.to_vec(),
            ));
        }
    }
    replace_spans(source, &replacements)
}

#[derive(Clone, Copy)]
enum Target {
    Slide,
    Notes,
}

#[derive(Clone)]
struct SourceSpan {
    start: usize,
    end: usize,
    close_start: usize,
    empty: bool,
    qname: String,
}

struct SourceFrame {
    start: usize,
    local: String,
    qname: String,
    ext_list: bool,
    target: Option<Target>,
}

struct SourceLayout {
    root: SourceSpan,
    root_close: Option<usize>,
    ext_list: Option<SourceSpan>,
    targets: [Option<SourceSpan>; 2],
    presentation_namespace: &'static str,
    strict: bool,
}

fn scan_source(source: &[u8]) -> Result<SourceLayout> {
    let mut reader = NsReader::from_reader(source);
    let mut stack: Vec<SourceFrame> = Vec::new();
    let mut root = None;
    let mut root_close = None;
    let mut ext_list = None;
    let mut targets = [None, None];
    let mut root_namespace = None;
    let mut nodes = 0usize;

    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("presentation-guide XML offset overflow"))?;
        let decoder = reader.decoder();
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(invalid("presentation-guide XML resource limit exceeded"));
                }
                let local = local_name(element.local_name().as_ref())?;
                let qname = std::str::from_utf8(element.name().as_ref())
                    .map_err(xml_error)?
                    .to_owned();
                let is_presentation = presentation_namespace(&namespace).is_some();
                if stack.is_empty() {
                    if root.is_some() || !is_presentation || local != "presentation" {
                        return Err(invalid("expected one PresentationML presentation root"));
                    }
                    root_namespace = presentation_namespace(&namespace);
                }
                let is_ext_list = stack.len() == 1 && is_presentation && local == "extLst";
                if is_ext_list && ext_list.is_some() {
                    return Err(invalid("duplicate presentation extLst"));
                }
                let target = if stack.last().is_some_and(|frame| frame.ext_list)
                    && is_presentation
                    && local == "ext"
                {
                    target_from_uri(element_uri(&element, decoder)?)
                } else {
                    None
                };
                stack.push(SourceFrame {
                    start,
                    local,
                    qname,
                    ext_list: is_ext_list,
                    target,
                });
            },
            Event::Empty(element) => {
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(invalid("presentation-guide XML node limit exceeded"));
                }
                let local = local_name(element.local_name().as_ref())?;
                let qname = std::str::from_utf8(element.name().as_ref())
                    .map_err(xml_error)?
                    .to_owned();
                let is_presentation = presentation_namespace(&namespace).is_some();
                if stack.is_empty() {
                    if root.is_some() || !is_presentation || local != "presentation" {
                        return Err(invalid("expected one PresentationML presentation root"));
                    }
                    root_namespace = presentation_namespace(&namespace);
                    root = Some(SourceSpan {
                        start,
                        end: reader.buffer_position() as usize,
                        close_start: start,
                        empty: true,
                        qname,
                    });
                } else if stack.len() == 1 && is_presentation && local == "extLst" {
                    if ext_list.is_some() {
                        return Err(invalid("duplicate presentation extLst"));
                    }
                    ext_list = Some(SourceSpan {
                        start,
                        end: reader.buffer_position() as usize,
                        close_start: start,
                        empty: true,
                        qname,
                    });
                } else if stack.len() == 2
                    && stack.last().is_some_and(|frame| frame.ext_list)
                    && is_presentation
                    && local == "ext"
                    && let Some(target) = target_from_uri(element_uri(&element, decoder)?)
                {
                    let index = target_index(target);
                    if targets[index].is_some() {
                        return Err(invalid("duplicate extended-guide extension"));
                    }
                    targets[index] = Some(SourceSpan {
                        start,
                        end: reader.buffer_position() as usize,
                        close_start: start,
                        empty: true,
                        qname,
                    });
                }
            },
            Event::End(element) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected presentation-guide closing element"))?;
                let local = local_name(element.local_name().as_ref())?;
                if frame.local != local {
                    return Err(invalid("mismatched presentation-guide closing element"));
                }
                let end = usize::try_from(reader.buffer_position())
                    .map_err(|_| invalid("presentation-guide XML offset overflow"))?;
                let span = SourceSpan {
                    start: frame.start,
                    end,
                    close_start: start,
                    empty: false,
                    qname: frame.qname,
                };
                if stack.is_empty() {
                    if root.is_some() {
                        return Err(invalid("multiple PresentationML presentation roots"));
                    }
                    root_close = Some(start);
                    root = Some(span);
                } else if frame.ext_list {
                    if ext_list.is_some() {
                        return Err(invalid("duplicate presentation extLst"));
                    }
                    ext_list = Some(span);
                } else if let Some(target) = frame.target {
                    let index = target_index(target);
                    if targets[index].is_some() {
                        return Err(invalid("duplicate extended-guide extension"));
                    }
                    targets[index] = Some(span);
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            Event::Decl(_) | Event::Text(_) | Event::CData(_) | Event::Comment(_) => {},
            Event::GeneralRef(_) => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated presentation-guide XML"));
    }
    let root = root.ok_or_else(|| invalid("missing presentation root"))?;
    let presentation_namespace =
        root_namespace.ok_or_else(|| invalid("missing PresentationML presentation namespace"))?;
    Ok(SourceLayout {
        strict: presentation_namespace == PS,
        root,
        root_close,
        ext_list,
        targets,
        presentation_namespace,
    })
}

fn local_name(name: &[u8]) -> Result<String> {
    let name = std::str::from_utf8(name).map_err(xml_error)?;
    Ok(name
        .rsplit_once(':')
        .map_or(name, |(_, local)| local)
        .to_owned())
}

fn presentation_namespace(namespace: &ResolveResult<'_>) -> Option<&'static str> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == P.as_bytes() => Some(P),
        ResolveResult::Bound(Namespace(value)) if *value == PS.as_bytes() => Some(PS),
        _ => None,
    }
}

fn element_uri(element: &BytesStart<'_>, decoder: Decoder) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == b"uri" {
            if value.is_some() {
                return Err(invalid("duplicate presentation-guide extension URI"));
            }
            value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(xml_error)?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

fn target_from_uri(uri: Option<String>) -> Option<Target> {
    match uri.as_deref() {
        Some(SLIDE_GUIDES_URI) => Some(Target::Slide),
        Some(NOTES_GUIDES_URI) => Some(Target::Notes),
        _ => None,
    }
}

fn target_index(target: Target) -> usize {
    match target {
        Target::Slide => 0,
        Target::Notes => 1,
    }
}

fn target_entries(
    value: &Guides,
    strict: bool,
    presentation_namespace: &str,
) -> Result<[Option<String>; 2]> {
    Ok([
        value
            .slide
            .as_ref()
            .map(|list| {
                extension_xml(
                    SLIDE_GUIDES_URI,
                    "sldGuideLst",
                    list,
                    strict,
                    presentation_namespace,
                )
            })
            .transpose()?,
        value
            .notes
            .as_ref()
            .map(|list| {
                extension_xml(
                    NOTES_GUIDES_URI,
                    "notesGuideLst",
                    list,
                    strict,
                    presentation_namespace,
                )
            })
            .transpose()?,
    ])
}

fn extension_xml(
    uri: &str,
    local: &str,
    list: &List,
    strict: bool,
    presentation_namespace: &str,
) -> Result<String> {
    let mut xml = String::new();
    write_list_extension(&mut xml, uri, local, list, strict)?;
    let declaration = format!("<p:ext xmlns:p=\"{presentation_namespace}\" ");
    Ok(xml.replacen("<p:ext ", &declaration, 1))
}

fn expand_empty(source: &[u8], span: &SourceSpan, body: &str) -> Result<Vec<u8>> {
    let replacement = expand_empty_bytes(source, span, body.as_bytes())?;
    Ok(replacement)
}

fn expand_empty_bytes(source: &[u8], span: &SourceSpan, body: &[u8]) -> Result<Vec<u8>> {
    let token = source[span.start..span.end].trim_ascii_end();
    let open = token
        .strip_suffix(b"/>")
        .ok_or_else(|| invalid("presentation-guide empty element has no self-close"))?;
    let mut replacement = Vec::with_capacity(
        open.len()
            .saturating_add(1)
            .saturating_add(body.len())
            .saturating_add(span.qname.len())
            .saturating_add(3),
    );
    replacement.extend_from_slice(open);
    replacement.push(b'>');
    replacement.extend_from_slice(body);
    replacement.extend_from_slice(b"</");
    replacement.extend_from_slice(span.qname.as_bytes());
    replacement.push(b'>');
    if replacement.len() > MAX_BYTES {
        return Err(invalid("patched presentation guides exceed 8 MiB"));
    }
    Ok(replacement)
}

fn replace_spans(source: &[u8], spans: &[(usize, usize, Vec<u8>)]) -> Result<Vec<u8>> {
    let mut ordered = spans.to_vec();
    ordered.sort_by_key(|(start, _, _)| *start);
    for pair in ordered.windows(2) {
        if pair[0].1 > pair[1].0 {
            return Err(invalid("overlapping presentation-guide patch ranges"));
        }
    }
    let mut output = source.to_vec();
    for (start, end, replacement) in ordered.iter().rev() {
        output.splice(*start..*end, replacement.iter().copied());
    }
    if output.len() > MAX_BYTES {
        return Err(invalid("patched presentation guides exceed 8 MiB"));
    }
    Ok(output)
}

fn insert_bytes(source: &[u8], offset: usize, replacement: &[u8]) -> Result<Vec<u8>> {
    if offset > source.len() {
        return Err(invalid(
            "presentation-guide insertion offset is out of bounds",
        ));
    }
    let mut output = Vec::with_capacity(source.len().saturating_add(replacement.len()));
    output.extend_from_slice(&source[..offset]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[offset..]);
    if output.len() > MAX_BYTES {
        return Err(invalid("patched presentation guides exceed 8 MiB"));
    }
    Ok(output)
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
                let text = litchi_ooxml_common::xml::decode_xml_reference(&reference)?;
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

fn parse_presentation(root: &Node) -> Result<Guides> {
    expect(root, &[P, PS], "presentation")?;
    let mut root_ext = None;
    for child in children(root)? {
        if is(child, &[P, PS], "extLst") && root_ext.replace(child).is_some() {
            return Err(invalid("duplicate presentation extLst"));
        }
    }
    let Some(root_ext) = root_ext else {
        return Ok(Guides::default());
    };
    let mut value = Guides::default();
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

fn parse_list(node: &Node) -> Result<List> {
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
    Ok(List {
        guides,
        extension_xml,
    })
}

fn parse_guide(node: &Node) -> Result<Guide> {
    let id = required_attr(node, "id")?
        .parse::<u32>()
        .map_err(|_| invalid("invalid extended guide ID"))?;
    let name = optional_attr(node, "name")?;
    if let Some(name) = &name {
        bounded_string(name)?;
    }
    let orientation = optional_attr(node, "orient")?
        .map(|value| match value.as_str() {
            "horz" => Ok(Orientation::Horizontal),
            "vert" => Ok(Orientation::Vertical),
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
    Ok(Guide {
        id,
        name,
        orientation,
        position,
        user_drawn,
        color,
        extension_xml,
    })
}

fn parse_color(node: &Node) -> Result<Color> {
    no_attributes(node)?;
    let colors = children(node)?;
    if colors.len() != 1 {
        return Err(invalid("extended guide clr requires one DrawingML color"));
    }
    let color = colors[0];
    let kind = match (color.namespace.as_str(), color.local.as_str()) {
        (A | AS, "scrgbClr") => ColorKind::ScRgb,
        (A | AS, "srgbClr") => ColorKind::Srgb,
        (A | AS, "hslClr") => ColorKind::Hsl,
        (A | AS, "sysClr") => ColorKind::System,
        (A | AS, "schemeClr") => ColorKind::Scheme,
        (A | AS, "prstClr") => ColorKind::Preset,
        _ => return Err(invalid("invalid extended guide DrawingML color")),
    };
    Ok(Color {
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

fn validate(value: &Guides) -> Result<()> {
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

fn validate_color(color: &Color) -> Result<()> {
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

fn parse_color_node(node: &Node) -> Result<ColorKind> {
    match (node.namespace.as_str(), node.local.as_str()) {
        (A | AS, "scrgbClr") => Ok(ColorKind::ScRgb),
        (A | AS, "srgbClr") => Ok(ColorKind::Srgb),
        (A | AS, "hslClr") => Ok(ColorKind::Hsl),
        (A | AS, "sysClr") => Ok(ColorKind::System),
        (A | AS, "schemeClr") => Ok(ColorKind::Scheme),
        (A | AS, "prstClr") => Ok(ColorKind::Preset),
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
    list: &List,
    strict: bool,
) -> Result<()> {
    write!(
        xml,
        "<p:ext uri=\"{uri}\"><p15:{local} xmlns:p15=\"{P15}\">"
    )
    .map_err(|error| Error::Xml(error.to_string()))?;
    for guide in &list.guides {
        write_guide(xml, guide, strict)?;
    }
    if let Some(extension) = &list.extension_xml {
        write_opaque(xml, extension, strict)?;
    }
    write!(xml, "</p15:{local}></p:ext>").map_err(|error| Error::Xml(error.to_string()))?;
    Ok(())
}

fn write_guide(xml: &mut String, guide: &Guide, strict: bool) -> Result<()> {
    write!(xml, "<p15:guide id=\"{}\"", guide.id).map_err(|error| Error::Xml(error.to_string()))?;
    if let Some(name) = &guide.name {
        xml.push_str(" name=\"");
        escape_attribute(xml, name);
        xml.push('"');
    }
    if let Some(orientation) = guide.orientation {
        xml.push_str(" orient=\"");
        xml.push_str(match orientation {
            Orientation::Horizontal => "horz",
            Orientation::Vertical => "vert",
        });
        xml.push('"');
    }
    if let Some(position) = guide.position {
        write!(xml, " pos=\"{position}\"").map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(user_drawn) = guide.user_drawn {
        write!(xml, " userDrawn=\"{}\"", if user_drawn { 1 } else { 0 })
            .map_err(|error| Error::Xml(error.to_string()))?;
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
    let (node_prefix, _) = split_qname(&node.qname)?;
    let mut required_prefixes = HashSet::new();
    if !node.namespace.is_empty() {
        required_prefixes.insert(node_prefix.to_owned());
    }
    for attribute in &node.attributes {
        if let Some((prefix, _)) = attribute.qname.split_once(':') {
            required_prefixes.insert(prefix.to_owned());
        }
    }
    xml.push('<');
    xml.push_str(&node.qname);
    for (prefix, uri) in &node.bindings {
        if !required_prefixes.contains(prefix) {
            continue;
        }
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

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn wrap(fragment: &str) -> String {
        format!(
            r#"<p:presentation xmlns:p="{P}" xmlns:a="{A}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{fragment}</p:presentation>"#
        )
    }

    #[test]
    fn round_trips_slide_notes_and_unknown_extensions_inertly() {
        let xml = wrap(&format!(
            r#"<p:extLst><p:ext uri="{SLIDE_GUIDES_URI}"><p15:sldGuideLst xmlns:p15="{P15}"><p15:guide id="7" name="Named" orient="vert" pos="-20" userDrawn="0"><p15:clr><a:schemeClr val="accent1"/></p15:clr><p:extLst><p:ext uri="urn:guide"><v:data xmlns:v="urn:vendor" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdNeverFetched" href="https://example.invalid/not-opened"/></p:ext></p:extLst></p15:guide><p:extLst><p:ext uri="urn:list"><v:list xmlns:v="urn:vendor"/></p:ext></p:extLst></p15:sldGuideLst></p:ext><p:ext uri="{NOTES_GUIDES_URI}"><p15:notesGuideLst xmlns:p15="{P15}"/></p:ext></p:extLst>"#
        ));
        let value = Guides::from_xml(xml.as_bytes()).unwrap();
        let guide = &value.slide.as_ref().unwrap().guides[0];
        assert_eq!(guide.name.as_deref(), Some("Named"));
        assert_eq!(guide.color.kind, ColorKind::Scheme);
        let opaque = std::str::from_utf8(guide.extension_xml.as_deref().unwrap()).unwrap();
        assert!(opaque.contains("rIdNeverFetched"));
        assert!(opaque.contains("https://example.invalid/not-opened"));
        assert!(value.notes.as_ref().unwrap().guides.is_empty());
        for strict in [false, true] {
            let written = value.to_xml(strict).unwrap();
            let again = Guides::from_xml(wrap(&written).as_bytes()).unwrap();
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
            assert!(Guides::from_xml(xml.as_bytes()).is_err(), "accepted {xml}");
        }
    }

    #[test]
    fn rejects_programmatic_limits_and_color_mismatch() {
        let color = Color {
            kind: ColorKind::Srgb,
            xml: format!(r#"<a:srgbClr xmlns:a="{A}" val="AABBCC"/>"#).into_bytes(),
        };
        let guide = Guide {
            id: 1,
            name: None,
            orientation: None,
            position: None,
            user_drawn: None,
            color: color.clone(),
            extension_xml: None,
        };
        let mut value = Guides {
            slide: Some(List {
                guides: vec![guide; MAX_GUIDES + 1],
                extension_xml: None,
            }),
            notes: None,
        };
        assert!(value.to_xml(false).is_err());
        value.slide.as_mut().unwrap().guides.truncate(1);
        value.slide.as_mut().unwrap().guides[0].color.kind = ColorKind::Scheme;
        assert!(value.to_xml(false).is_err());
    }
}
