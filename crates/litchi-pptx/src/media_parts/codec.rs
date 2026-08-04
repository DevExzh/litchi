use super::*;

#[derive(Clone)]
struct Attribute {
    namespace: String,
    prefix: String,
    name: String,
    value: String,
}

#[derive(Default)]
struct NamespaceContext {
    parent: Option<Arc<Self>>,
    declarations: Vec<(String, String)>,
}

#[derive(Clone)]
struct Node {
    namespace: String,
    prefix: String,
    name: String,
    attributes: Vec<Attribute>,
    children: Vec<Node>,
    /// Direct text plus UTF-8 boundaries immediately before each child.
    text: String,
    text_ends: Vec<usize>,
    namespace_context: Arc<NamespaceContext>,
    declares_namespaces: bool,
}

impl ExtensionList {
    /// Parse one transitional PresentationML `p:extLst` fragment.
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_MEDIA_EXTENSION_XML_BYTES {
            return Err(limit("media extension-list XML bytes"));
        }
        let root = parse_document(xml)?;
        Self::from_node(&root)
    }

    /// Borrow the self-contained canonical XML fragment.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.xml
    }

    fn from_node(node: &Node) -> Result<Self> {
        require_node(node, PML, "extLst")?;
        let xml = canonical_fragment(node)?;
        let xml = String::from_utf8(xml)
            .map_err(|_| invalid("canonical media extension-list XML is not UTF-8"))?;
        Ok(Self {
            xml: xml.into_boxed_str(),
        })
    }
}

/// Parses all audio/video `p:pic` elements from a complete Slide part.
pub fn parse(xml: &[u8]) -> Result<List> {
    let root = parse_document(xml)?;
    let conformance = conformance(&root)?;
    let mut pictures = Vec::new();
    collect_pictures(&root, conformance, &mut pictures)?;
    let value = List { pictures };
    validate_value(&value, false)?;
    Ok(value)
}

fn collect_pictures(
    node: &Node,
    conformance: Conformance,
    output: &mut Vec<Picture>,
) -> Result<()> {
    if node.namespace == conformance.pml()
        && node.name == "pic"
        && let Some(picture) = parse_picture(node, conformance)?
    {
        if output.len() == MAX_MEDIA {
            return Err(limit("media count"));
        }
        output.push(picture);
    }
    for child in &node.children {
        collect_pictures(child, conformance, output)?;
    }
    Ok(())
}

fn parse_picture(node: &Node, conformance: Conformance) -> Result<Option<Picture>> {
    let Some(nv_pic) = one_child(node, conformance.pml(), "nvPicPr")? else {
        return Ok(None);
    };
    let Some(nv_pr) = one_child(nv_pic, conformance.pml(), "nvPr")? else {
        return Ok(None);
    };
    let audio = one_child(nv_pr, conformance.dml(), "audioFile")?;
    let video = one_child(nv_pr, conformance.dml(), "videoFile")?;
    let (kind, media) = match (audio, video) {
        (Some(_), Some(_)) => {
            return Err(invalid(
                "media picture contains both audioFile and videoFile",
            ));
        },
        (Some(value), None) => (Kind::Audio, value),
        (None, Some(value)) => (Kind::Video, value),
        (None, None) => return Ok(None),
    };
    leaf(media, "media file")?;
    let relationship_id = required(media, conformance.rel(), "link")?.to_owned();
    no_attributes(media, &[(conformance.rel(), "link")])?;
    let c_nv_pr = required_child(nv_pic, conformance.pml(), "cNvPr")?;
    let shape_id = required(c_nv_pr, "", "id")?
        .parse()
        .map_err(|_| invalid("invalid media shape id"))?;
    let name = optional(c_nv_pr, "", "name").unwrap_or_default().to_owned();
    let poster = parse_poster(node, conformance)?;
    let transform = parse_transform(node, conformance)?;
    let office_extension = find_office_media(nv_pr, conformance)?
        .map(parse_office_media)
        .transpose()?;
    Ok(Some(Picture {
        shape_id,
        name,
        kind,
        relationship_id,
        resource: None,
        poster,
        transform,
        office_extension,
    }))
}

fn parse_poster(node: &Node, conformance: Conformance) -> Result<Option<Poster>> {
    let Some(fill) = one_child(node, conformance.pml(), "blipFill")? else {
        return Ok(None);
    };
    let Some(blip) = one_child(fill, conformance.dml(), "blip")? else {
        return Ok(None);
    };
    let relationship_id = required(blip, conformance.rel(), "embed")?.to_owned();
    Ok(Some(Poster {
        relationship_id,
        resource: None,
    }))
}

fn parse_transform(node: &Node, conformance: Conformance) -> Result<Option<Transform>> {
    let Some(properties) = one_child(node, conformance.pml(), "spPr")? else {
        return Ok(None);
    };
    let Some(transform) = one_child(properties, conformance.dml(), "xfrm")? else {
        return Ok(None);
    };
    let offset = required_child(transform, conformance.dml(), "off")?;
    let extent = required_child(transform, conformance.dml(), "ext")?;
    leaf(offset, "transform offset")?;
    leaf(extent, "transform extent")?;
    Ok(Some(Transform::new(
        parse_coordinate(required(offset, "", "x")?, "x")?,
        parse_coordinate(required(offset, "", "y")?, "y")?,
        parse_extent(required(extent, "", "cx")?, "width")?,
        parse_extent(required(extent, "", "cy")?, "height")?,
    )))
}

fn find_office_media(node: &Node, conformance: Conformance) -> Result<Option<&Node>> {
    let mut found = None;
    for list in node
        .children
        .iter()
        .filter(|child| child.namespace == conformance.pml() && child.name == "extLst")
    {
        for extension in list
            .children
            .iter()
            .filter(|child| child.namespace == conformance.pml() && child.name == "ext")
        {
            for media in extension
                .children
                .iter()
                .filter(|child| child.namespace == P14 && child.name == "media")
            {
                if found.replace(media).is_some() {
                    return Err(invalid("media picture has multiple p14:media extensions"));
                }
            }
        }
    }
    Ok(found)
}

fn parse_office_media(node: &Node) -> Result<Extension> {
    whitespace(node)?;
    // The Office 2010 p14 schema imports the transitional relationships
    // namespace even when the containing presentation uses Strict namespaces.
    let embed_relationship_id = optional(node, REL, "embed").map(str::to_owned);
    let link_relationship_id = optional(node, REL, "link").map(str::to_owned);
    no_attributes(node, &[(REL, "embed"), (REL, "link")])?;
    if embed_relationship_id.is_none() && link_relationship_id.is_none() {
        return Err(invalid("p14:media requires r:embed or r:link"));
    }
    let mut trim_node = None;
    let mut fade_node = None;
    let mut bookmarks_node = None;
    let mut extensions = None;
    let mut stage = 0u8;
    for child in &node.children {
        let child_stage = match (child.namespace.as_str(), child.name.as_str()) {
            (P14, "trim") => {
                if trim_node.replace(child).is_some() {
                    return Err(invalid("p14:media has multiple trim children"));
                }
                1
            },
            (P14, "fade") => {
                if fade_node.replace(child).is_some() {
                    return Err(invalid("p14:media has multiple fade children"));
                }
                2
            },
            (P14, "bmkLst") => {
                if bookmarks_node.replace(child).is_some() {
                    return Err(invalid("p14:media has multiple bmkLst children"));
                }
                3
            },
            // The p14 schema imports transitional PresentationML even when
            // the containing slide uses the Strict namespace dialect.
            (PML, "extLst") => {
                if extensions
                    .replace(ExtensionList::from_node(child)?)
                    .is_some()
                {
                    return Err(invalid("p14:media has multiple extLst children"));
                }
                4
            },
            _ => {
                return Err(invalid(format!(
                    "unsupported p14:media child '{}'",
                    child.name
                )));
            },
        };
        if child_stage < stage {
            return Err(invalid(
                "p14:media children are outside the schema-defined order",
            ));
        }
        stage = child_stage;
    }
    let trim = trim_node.map(parse_trim).transpose()?;
    let fade = fade_node.map(parse_fade).transpose()?;
    let mut bookmarks = Vec::new();
    if let Some(list) = bookmarks_node {
        whitespace(list)?;
        no_attributes(list, &[])?;
        if list.children.len() > MAX_BOOKMARKS {
            return Err(limit("bookmark count"));
        }
        for child in &list.children {
            require_node(child, P14, "bmk")?;
            leaf(child, "media bookmark")?;
            bookmarks.push(Bookmark {
                name: optional(child, "", "name").map(str::to_owned),
                time: optional(child, "", "time").map(parse_time).transpose()?,
            });
            no_attributes(child, &[("", "name"), ("", "time")])?;
        }
    }
    Ok(Extension {
        embed_relationship_id,
        link_relationship_id,
        trim,
        fade,
        bookmarks,
        extensions,
    })
}

fn parse_trim(node: &Node) -> Result<Trim> {
    leaf(node, "media trim")?;
    no_attributes(node, &[("", "st"), ("", "end")])?;
    Ok(Trim {
        start: optional(node, "", "st").map(parse_time).transpose()?,
        end: optional(node, "", "end").map(parse_time).transpose()?,
    })
}

fn parse_fade(node: &Node) -> Result<Fade> {
    leaf(node, "media fade")?;
    no_attributes(node, &[("", "in"), ("", "out")])?;
    Ok(Fade {
        fade_in: optional(node, "", "in").map(parse_time).transpose()?,
        fade_out: optional(node, "", "out").map(parse_time).transpose()?,
    })
}

/// Deterministically serializes self-contained `p:pic` fragments.
pub fn write_pictures(value: &List, conformance: Conformance) -> Result<Vec<u8>> {
    validate_value(value, false)?;
    let mut output = BoundedXml::new();
    for picture in &value.pictures {
        write_picture(&mut output, picture, conformance)?;
    }
    Ok(output.finish())
}

pub(crate) fn write_picture(
    output: &mut BoundedXml,
    picture: &Picture,
    conformance: Conformance,
) -> Result<()> {
    output.write(b"<p:pic xmlns:p=\"")?;
    output.escape(conformance.pml())?;
    output.write(b"\" xmlns:a=\"")?;
    output.escape(conformance.dml())?;
    output.write(b"\" xmlns:r=\"")?;
    output.escape(conformance.rel())?;
    if picture.office_extension.is_some() {
        output.write(b"\" xmlns:p14=\"")?;
        output.escape(P14)?;
    }
    output.write(b"\"><p:nvPicPr><p:cNvPr")?;
    output.attr("id", &picture.shape_id.to_string())?;
    output.attr("name", &picture.name)?;
    output.write(b"/><p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr>")?;
    output.write(match picture.kind {
        Kind::Audio => b"<a:audioFile".as_slice(),
        Kind::Video => b"<a:videoFile".as_slice(),
    })?;
    output.attr("r:link", &picture.relationship_id)?;
    output.write(b"/>")?;
    if let Some(extension) = &picture.office_extension {
        write_office_extension(output, extension)?;
    }
    output.write(b"</p:nvPr></p:nvPicPr><p:blipFill>")?;
    if let Some(poster) = &picture.poster {
        output.write(b"<a:blip")?;
        output.attr("r:embed", &poster.relationship_id)?;
        output.write(b"/><a:stretch><a:fillRect/></a:stretch>")?;
    }
    output.write(b"</p:blipFill><p:spPr>")?;
    if let Some(transform) = &picture.transform {
        output.write(b"<a:xfrm><a:off")?;
        output.attr("x", &transform.x.to_string())?;
        output.attr("y", &transform.y.to_string())?;
        output.write(b"/><a:ext")?;
        output.attr("cx", &transform.width.to_string())?;
        output.attr("cy", &transform.height.to_string())?;
        output.write(b"/></a:xfrm>")?;
    }
    output.write(b"<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>")?;
    Ok(())
}

fn write_office_extension(output: &mut BoundedXml, value: &Extension) -> Result<()> {
    output.write(b"<p:extLst><p:ext uri=\"")?;
    output.escape(MEDIA_EXTENSION_URI)?;
    output.write(b"\"><p14:media xmlns:r=\"")?;
    output.escape(REL)?;
    output.push(b'"')?;
    if let Some(id) = &value.embed_relationship_id {
        output.attr("r:embed", id)?;
    }
    if let Some(id) = &value.link_relationship_id {
        output.attr("r:link", id)?;
    }
    if value.trim.is_none()
        && value.fade.is_none()
        && value.bookmarks.is_empty()
        && value.extensions.is_none()
    {
        output.write(b"/></p:ext></p:extLst>")?;
        return Ok(());
    }
    output.push(b'>')?;
    if let Some(trim) = &value.trim {
        output.write(b"<p14:trim")?;
        if let Some(start) = &trim.start {
            output.attr("st", start.as_str())?;
        }
        if let Some(end) = &trim.end {
            output.attr("end", end.as_str())?;
        }
        output.write(b"/>")?;
    }
    if let Some(fade) = &value.fade {
        output.write(b"<p14:fade")?;
        if let Some(fade_in) = &fade.fade_in {
            output.attr("in", fade_in.as_str())?;
        }
        if let Some(fade_out) = &fade.fade_out {
            output.attr("out", fade_out.as_str())?;
        }
        output.write(b"/>")?;
    }
    if !value.bookmarks.is_empty() {
        output.write(b"<p14:bmkLst>")?;
        for bookmark in &value.bookmarks {
            output.write(b"<p14:bmk")?;
            if let Some(v) = &bookmark.name {
                output.attr("name", v)?;
            }
            if let Some(v) = &bookmark.time {
                output.attr("time", v.as_str())?;
            }
            output.write(b"/>")?;
        }
        output.write(b"</p14:bmkLst>")?;
    }
    if let Some(extensions) = &value.extensions {
        output.write(extensions.as_str().as_bytes())?;
    }
    output.write(b"</p14:media></p:ext></p:extLst>")?;
    Ok(())
}

/// Loads media pictures and validates their complete internal OPC resource graph.
fn parse_document(xml: &[u8]) -> Result<Node> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("input XML bytes"));
    }
    let mut caps = MceCapabilities::ooxml_baseline();
    caps.understand_namespace(P14);
    for namespace in [PML, STRICT_PML] {
        caps.preserve_extension_element(ExpandedName {
            namespace: namespace.to_owned(),
            local_name: "ext".to_owned(),
        });
    }
    let limits = MceLimits {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        max_depth: MAX_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let processed = process_markup_compatibility(xml, &caps, &limits)?.xml;
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    let base_namespace_context = Arc::new(NamespaceContext::default());
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(limit("XML structure"));
                }
                let empty = matches!(&event, Event::Empty(_));
                let parent_namespace_context = stack
                    .last()
                    .map_or(&base_namespace_context, |node| &node.namespace_context);
                let node = make_node(
                    &reader,
                    element,
                    reader.decoder(),
                    &mut strings,
                    parent_namespace_context,
                )?;
                if empty {
                    attach(node, &mut stack, &mut root)?;
                } else {
                    stack.push(node);
                }
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                attach(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                add_strings(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("text outside slide root"));
                }
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value = reference
                    .resolve_char_ref()
                    .map_err(xml_error)?
                    .map(|v| v.to_string())
                    .or_else(|| match name.as_ref() {
                        "amp" => Some("&".into()),
                        "lt" => Some("<".into()),
                        "gt" => Some(">".into()),
                        "apos" => Some("'".into()),
                        "quot" => Some("\"".into()),
                        _ => None,
                    })
                    .ok_or_else(|| invalid("custom XML entity is rejected"))?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else {
                    return Err(invalid("entity outside slide root"));
                }
            },
            Event::CData(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                add_strings(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("CDATA outside slide root"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated slide XML"));
    }
    root.ok_or_else(|| invalid("missing slide root"))
}

fn make_node(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    strings: &mut usize,
    parent_namespace_context: &Arc<NamespaceContext>,
) -> Result<Node> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let prefix = element
        .name()
        .prefix()
        .map(|prefix| std::str::from_utf8(prefix.as_ref()).map(str::to_owned))
        .transpose()
        .map_err(xml_error)?
        .unwrap_or_default();
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    add_strings(strings, namespace.len() + prefix.len() + name.len())?;
    let mut attributes = Vec::new();
    let mut namespace_declarations = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if let Some(declaration) = item.key.as_namespace_binding() {
            let prefix = match declaration {
                PrefixDeclaration::Default => String::new(),
                PrefixDeclaration::Named(prefix) => {
                    std::str::from_utf8(prefix).map_err(xml_error)?.to_owned()
                },
            };
            add_strings(strings, prefix.len() + value.len())?;
            namespace_declarations.push((prefix, value));
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let prefix = item
            .key
            .prefix()
            .map(|prefix| std::str::from_utf8(prefix.as_ref()).map(str::to_owned))
            .transpose()
            .map_err(xml_error)?
            .unwrap_or_default();
        let name = std::str::from_utf8(local.as_ref())
            .map_err(xml_error)?
            .to_owned();
        add_strings(
            strings,
            namespace.len() + prefix.len() + name.len() + value.len(),
        )?;
        if attributes
            .iter()
            .any(|attribute: &Attribute| attribute.namespace == namespace && attribute.name == name)
        {
            return Err(invalid("duplicate expanded XML attribute"));
        }
        attributes.push(Attribute {
            namespace,
            prefix,
            name,
            value,
        });
    }
    let declares_namespaces = !namespace_declarations.is_empty();
    let namespace_context = if declares_namespaces {
        Arc::new(NamespaceContext {
            parent: Some(Arc::clone(parent_namespace_context)),
            declarations: namespace_declarations,
        })
    } else {
        Arc::clone(parent_namespace_context)
    };
    Ok(Node {
        namespace,
        prefix,
        name,
        attributes,
        children: Vec::new(),
        text: String::new(),
        text_ends: Vec::new(),
        namespace_context,
        declares_namespaces,
    })
}

fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.text_ends.push(parent.text.len());
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}
pub(crate) fn document_conformance(xml: &[u8]) -> Result<Conformance> {
    let root = parse_document(xml)?;
    conformance(&root)
}

fn conformance(root: &Node) -> Result<Conformance> {
    crate_conformance(root)
}
fn crate_conformance(root: &Node) -> Result<Conformance> {
    if root.name != "sld" {
        return Err(invalid("expected Slide root"));
    }
    match root.namespace.as_str() {
        PML => Ok(Conformance::Transitional),
        STRICT_PML => Ok(Conformance::Strict),
        _ => Err(invalid("unsupported Slide namespace")),
    }
}
fn one_child<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<Option<&'a Node>> {
    let mut values = node
        .children
        .iter()
        .filter(|child| child.namespace == namespace && child.name == name);
    let value = values.next();
    if values.next().is_some() {
        Err(invalid(format!(
            "{} has multiple {name} children",
            node.name
        )))
    } else {
        Ok(value)
    }
}
fn required_child<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a Node> {
    one_child(node, namespace, name)?
        .ok_or_else(|| invalid(format!("{} is missing {name}", node.name)))
}
fn require_node(node: &Node, namespace: &str, name: &str) -> Result<()> {
    if node.namespace == namespace && node.name == name {
        Ok(())
    } else {
        Err(invalid(format!("expected {name}, got {}", node.name)))
    }
}
fn optional<'a>(node: &'a Node, namespace: &str, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace == namespace && attribute.name == name)
        .map(|attribute| attribute.value.as_str())
}
fn required<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a str> {
    optional(node, namespace, name)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid(format!("{} is missing attribute '{name}'", node.name)))
}
fn no_attributes(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    if let Some(attribute) = node.attributes.iter().find(|attribute| {
        !allowed.contains(&(attribute.namespace.as_str(), attribute.name.as_str()))
    }) {
        Err(invalid(format!(
            "unexpected attribute '{}' on {}",
            attribute.name, node.name
        )))
    } else {
        Ok(())
    }
}
fn whitespace(node: &Node) -> Result<()> {
    if node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("unexpected text in {}", node.name)))
    }
}
fn leaf(node: &Node, name: &str) -> Result<()> {
    whitespace(node)?;
    if node.children.is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("{name} must not contain child elements")))
    }
}
fn parse_coordinate(value: &str, name: &str) -> Result<Coordinate> {
    Coordinate::parse(value).map_err(|error| coordinate_error(error, name))
}

fn parse_extent(value: &str, name: &str) -> Result<Extent> {
    Extent::parse(value).map_err(|error| coordinate_error(error, name))
}

pub(crate) fn validate_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return Err(invalid("relationship ID cannot be empty"));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        Err(invalid(format!("invalid relationship ID '{value}'")))
    } else {
        Ok(())
    }
}
pub(crate) fn bounded(value: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(limit("string bytes"))
    }
}
fn add_strings(total: &mut usize, size: usize) -> Result<()> {
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("XML string bytes"))?;
    if *total > MAX_STRING_BYTES {
        Err(limit("XML string bytes"))
    } else {
        Ok(())
    }
}
fn resolved(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn canonical_fragment(node: &Node) -> Result<Vec<u8>> {
    let mut used_prefixes = BTreeSet::new();
    collect_used_prefixes(node, &mut used_prefixes);
    let mut output = BoundedXml::with_limit(MAX_MEDIA_EXTENSION_XML_BYTES);
    write_canonical_node(&mut output, node, &used_prefixes, true)?;
    Ok(output.finish())
}

fn collect_used_prefixes(node: &Node, prefixes: &mut BTreeSet<String>) {
    if !node.prefix.is_empty() {
        prefixes.insert(node.prefix.clone());
    } else if !node.namespace.is_empty() {
        prefixes.insert(String::new());
    }
    for attribute in &node.attributes {
        if !attribute.prefix.is_empty() {
            prefixes.insert(attribute.prefix.clone());
        }
        for token in attribute.value.split(|character: char| {
            character.is_whitespace()
                || matches!(
                    character,
                    ',' | ';' | '(' | ')' | '[' | ']' | '{' | '}' | '/' | '\\' | '"' | '\''
                )
        }) {
            if let Some((prefix, _)) = token.split_once(':')
                && !prefix.is_empty()
            {
                prefixes.insert(prefix.to_owned());
            }
        }
        if attribute.namespace == MCE_NAMESPACE
            && matches!(attribute.name.as_str(), "Ignorable" | "Requires")
        {
            prefixes.extend(
                attribute
                    .value
                    .split_whitespace()
                    .filter(|prefix| !prefix.is_empty())
                    .map(str::to_owned),
            );
        }
    }
    for child in &node.children {
        collect_used_prefixes(child, prefixes);
    }
}

fn write_canonical_node(
    output: &mut BoundedXml,
    node: &Node,
    used_prefixes: &BTreeSet<String>,
    root: bool,
) -> Result<()> {
    if node.text_ends.len() != node.children.len() {
        return Err(invalid("invalid XML text-segment state"));
    }

    output.push(b'<')?;
    write_original_name(output, &node.prefix, &node.name)?;
    if root {
        let mut bindings = effective_namespace_bindings(&node.namespace_context)?;
        bindings.retain(|prefix, _| used_prefixes.contains(prefix));
        bindings.entry(String::new()).or_default();
        for (prefix, namespace) in &bindings {
            if prefix != "xml" {
                write_namespace_declaration(output, prefix, namespace)?;
            }
        }
    } else if node.declares_namespaces {
        let mut declarations: Vec<_> = node.namespace_context.declarations.iter().collect();
        declarations.sort_unstable_by(|left, right| left.0.cmp(&right.0));
        for (prefix, namespace) in declarations {
            if used_prefixes.contains(prefix) {
                write_namespace_declaration(output, prefix, namespace)?;
            }
        }
    }

    let mut attributes: Vec<_> = node.attributes.iter().collect();
    attributes.sort_unstable_by(|left, right| {
        (&left.namespace, &left.name, &left.prefix).cmp(&(
            &right.namespace,
            &right.name,
            &right.prefix,
        ))
    });
    for attribute in attributes {
        output.push(b' ')?;
        write_original_name(output, &attribute.prefix, &attribute.name)?;
        output.write(b"=\"")?;
        output.escape(&attribute.value)?;
        output.push(b'"')?;
    }

    if node.children.is_empty() && node.text.is_empty() {
        output.write(b"/>")?;
        return Ok(());
    }
    output.push(b'>')?;
    let mut text_start = 0usize;
    for (index, child) in node.children.iter().enumerate() {
        let text_end = *node
            .text_ends
            .get(index)
            .ok_or_else(|| invalid("invalid XML text-segment state"))?;
        let text = node
            .text
            .get(text_start..text_end)
            .ok_or_else(|| invalid("invalid XML text-segment state"))?;
        output.escape(text)?;
        write_canonical_node(output, child, used_prefixes, false)?;
        text_start = text_end;
    }
    let text = node
        .text
        .get(text_start..)
        .ok_or_else(|| invalid("invalid XML text-segment state"))?;
    output.escape(text)?;
    output.write(b"</")?;
    write_original_name(output, &node.prefix, &node.name)?;
    output.push(b'>')?;
    Ok(())
}

fn effective_namespace_bindings(
    context: &Arc<NamespaceContext>,
) -> Result<BTreeMap<String, String>> {
    let mut contexts = Vec::new();
    let mut current = Some(context.as_ref());
    while let Some(context) = current {
        if contexts.len() == MAX_DEPTH {
            return Err(limit("media extension namespace depth"));
        }
        contexts.push(context);
        current = context.parent.as_deref();
    }

    let mut bindings = BTreeMap::new();
    for context in contexts.into_iter().rev() {
        for (prefix, namespace) in &context.declarations {
            if namespace.is_empty() {
                bindings.remove(prefix);
            } else {
                bindings.insert(prefix.clone(), namespace.clone());
            }
        }
    }
    Ok(bindings)
}

fn write_namespace_declaration(
    output: &mut BoundedXml,
    prefix: &str,
    namespace: &str,
) -> Result<()> {
    if prefix.is_empty() {
        output.write(b" xmlns=\"")?;
    } else {
        output.write(b" xmlns:")?;
        output.write(prefix.as_bytes())?;
        output.write(b"=\"")?;
    }
    output.escape(namespace)?;
    output.push(b'"')
}

fn write_original_name(output: &mut BoundedXml, prefix: &str, name: &str) -> Result<()> {
    if !prefix.is_empty() {
        output.write(prefix.as_bytes())?;
        output.push(b':')?;
    }
    output.write(name.as_bytes())
}

pub(crate) struct BoundedXml {
    pub(super) bytes: Vec<u8>,
    limit: usize,
}

impl BoundedXml {
    pub(super) fn new() -> Self {
        Self::with_limit(MAX_XML_BYTES)
    }

    pub(super) fn with_limit(limit: usize) -> Self {
        Self {
            bytes: Vec::new(),
            limit,
        }
    }

    fn finish(self) -> Vec<u8> {
        self.bytes
    }

    fn write(&mut self, value: &[u8]) -> Result<()> {
        self.reserve(value.len())?;
        self.bytes.extend_from_slice(value);
        Ok(())
    }

    fn push(&mut self, value: u8) -> Result<()> {
        self.reserve(1)?;
        self.bytes.push(value);
        Ok(())
    }

    fn escape(&mut self, value: &str) -> Result<()> {
        let escaped = escaped_len(value).ok_or_else(|| output_limit(self.limit))?;
        self.reserve(escaped)?;
        write_escaped(&mut self.bytes, value);
        Ok(())
    }

    fn attr(&mut self, name: &str, value: &str) -> Result<()> {
        let escaped = escaped_len(value).ok_or_else(|| output_limit(self.limit))?;
        let fixed = 4usize
            .checked_add(name.len())
            .ok_or_else(|| output_limit(self.limit))?;
        let required = fixed
            .checked_add(escaped)
            .ok_or_else(|| output_limit(self.limit))?;
        self.reserve(required)?;
        self.bytes.push(b' ');
        self.bytes.extend_from_slice(name.as_bytes());
        self.bytes.extend_from_slice(b"=\"");
        write_escaped(&mut self.bytes, value);
        self.bytes.push(b'\"');
        Ok(())
    }

    fn reserve(&mut self, additional: usize) -> Result<()> {
        let next = self
            .bytes
            .len()
            .checked_add(additional)
            .ok_or_else(|| output_limit(self.limit))?;
        if next > self.limit {
            return Err(output_limit(self.limit));
        }
        let spare = self.bytes.capacity().saturating_sub(self.bytes.len());
        if additional > spare {
            self.bytes
                .try_reserve_exact(additional)
                .map_err(|source| Error::Allocation {
                    resource: "slide media serialized XML",
                    source,
                })?;
        }
        Ok(())
    }
}

fn escaped_len(value: &str) -> Option<usize> {
    value.chars().try_fold(0usize, |total, character| {
        let bytes = match character {
            '&' => 5,
            '<' => 4,
            '>' => 4,
            '"' => 6,
            '\t' => 5,
            '\n' | '\r' => 6,
            _ => character.len_utf8(),
        };
        total.checked_add(bytes)
    })
}

fn write_escaped(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}
pub(crate) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
