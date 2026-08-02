//! ODF 1.3 `style:master-page` schema validation and lossless mutation.

use std::collections::HashSet;

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

use crate::odt::{HeaderFooterKind, MasterPage, MasterPageChild, MasterPageChildKind};

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const DR3D: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
const ANIM: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:animation:1.0";
const PRESENTATION: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_FRAGMENT_BYTES: usize = 16 * 1_048_576;
const MAX_TOTAL_CHILD_BYTES: usize = 32 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_EVENTS: usize = 1_000_000;
const MAX_MASTER_PAGES: usize = 65_536;
const MAX_CHILDREN: usize = 65_536;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[derive(Clone)]
struct ExpandedName {
    namespace: Option<Vec<u8>>,
    local: Vec<u8>,
}

struct ActiveChild {
    kind: MasterPageChildKind,
    start: usize,
    depth: usize,
}

struct ActiveMaster {
    page: usize,
    depth: usize,
    last_rank: Option<u8>,
    seen: [bool; 11],
    child: Option<ActiveChild>,
    total_child_bytes: usize,
}

fn expanded(reader: &NsReader<&[u8]>, name: quick_xml::name::QName<'_>) -> ExpandedName {
    let (namespace, local) = reader.resolver().resolve_element(name);
    ExpandedName {
        namespace: match namespace {
            ResolveResult::Bound(Namespace(value)) => Some(value.to_vec()),
            _ => None,
        },
        local: local.as_ref().to_vec(),
    }
}

fn is_name(value: &ExpandedName, namespace: &[u8], local: &[u8]) -> bool {
    value.namespace.as_deref() == Some(namespace) && value.local == local
}

fn classify(value: &ExpandedName) -> Option<MasterPageChildKind> {
    if value.namespace.as_deref() == Some(STYLE) {
        return Some(match value.local.as_slice() {
            b"header" => MasterPageChildKind::HeaderFooter(HeaderFooterKind::Header),
            b"header-left" => MasterPageChildKind::HeaderFooter(HeaderFooterKind::HeaderLeft),
            b"header-first" => MasterPageChildKind::HeaderFooter(HeaderFooterKind::HeaderFirst),
            b"footer" => MasterPageChildKind::HeaderFooter(HeaderFooterKind::Footer),
            b"footer-left" => MasterPageChildKind::HeaderFooter(HeaderFooterKind::FooterLeft),
            b"footer-first" => MasterPageChildKind::HeaderFooter(HeaderFooterKind::FooterFirst),
            _ => return None,
        });
    }
    if is_name(value, DRAW, b"layer-set") {
        return Some(MasterPageChildKind::LayerSet);
    }
    if is_name(value, OFFICE, b"forms") {
        return Some(MasterPageChildKind::Forms);
    }
    if is_name(value, PRESENTATION, b"notes") {
        return Some(MasterPageChildKind::Notes);
    }
    if value.namespace.as_deref() == Some(ANIM)
        && matches!(
            value.local.as_slice(),
            b"par"
                | b"seq"
                | b"iterate"
                | b"animate"
                | b"set"
                | b"animateMotion"
                | b"animateColor"
                | b"animateTransform"
                | b"transitionFilter"
                | b"audio"
                | b"command"
        )
    {
        return Some(MasterPageChildKind::Animation);
    }
    if value.namespace.as_deref() == Some(DRAW)
        && matches!(
            value.local.as_slice(),
            b"a" | b"rect"
                | b"line"
                | b"polyline"
                | b"polygon"
                | b"regular-polygon"
                | b"path"
                | b"circle"
                | b"ellipse"
                | b"g"
                | b"page-thumbnail"
                | b"frame"
                | b"measure"
                | b"caption"
                | b"connector"
                | b"control"
                | b"custom-shape"
        )
        || is_name(value, DR3D, b"scene")
    {
        return Some(MasterPageChildKind::Shape);
    }
    None
}

fn register_child(active: &mut ActiveMaster, kind: MasterPageChildKind) -> Result<()> {
    if active.child.is_some() {
        return Err(invalid("master-page direct children overlap"));
    }
    let rank = kind.order();
    if active.last_rank.is_some_and(|previous| rank < previous) {
        return Err(invalid(
            "style:master-page children are out of ODF 1.3 schema order",
        ));
    }
    if kind != MasterPageChildKind::Shape && active.seen[usize::from(rank)] {
        return Err(invalid(format!("duplicate {:?} master-page child", kind)));
    }
    active.seen[usize::from(rank)] = true;
    active.last_rank = Some(rank);
    Ok(())
}

fn push_child(
    xml: &str,
    pages: &mut [MasterPage],
    active: &mut ActiveMaster,
    kind: MasterPageChildKind,
    start: usize,
    end: usize,
) -> Result<()> {
    let size = end
        .checked_sub(start)
        .ok_or_else(|| invalid("invalid master-page child span"))?;
    if size > MAX_FRAGMENT_BYTES {
        return Err(invalid("master-page child exceeds 16 MiB"));
    }
    active.total_child_bytes = active
        .total_child_bytes
        .checked_add(size)
        .ok_or_else(|| invalid("master-page child size overflows"))?;
    if active.total_child_bytes > MAX_TOTAL_CHILD_BYTES {
        return Err(invalid("master-page children exceed 32 MiB"));
    }
    let children = &mut pages[active.page].children;
    if children.len() >= MAX_CHILDREN {
        return Err(invalid("master-page has too many children"));
    }
    children.push(MasterPageChild {
        kind,
        xml: xml[start..end].to_string(),
    });
    Ok(())
}

fn required_style_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<String> {
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| invalid(format!("invalid master-page attribute: {error}")))?;
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == STYLE)
            && name.as_ref() == local
        {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| Error::XmlError(error.to_string()))?
                .into_owned();
            if value.is_empty() {
                return Err(invalid(format!(
                    "style:{} must not be empty",
                    String::from_utf8_lossy(local)
                )));
            }
            return Ok(value);
        }
    }
    Err(invalid(format!(
        "style:master-page is missing style:{}",
        String::from_utf8_lossy(local)
    )))
}

fn validate_header(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    page: &MasterPage,
) -> Result<()> {
    let name = required_style_attr(reader, element, b"name")?;
    let layout = required_style_attr(reader, element, b"page-layout-name")?;
    if name != page.name || page.page_layout_name.as_deref() != Some(layout.as_str()) {
        return Err(invalid(
            "master-page typed attributes disagree with lossless XML",
        ));
    }
    Ok(())
}

pub(crate) fn validate_master_page_schema(xml: &str, pages: &mut [MasterPage]) -> Result<()> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("master-page XML exceeds 64 MiB"));
    }
    if pages.len() > MAX_MASTER_PAGES {
        return Err(invalid("styles XML has too many master pages"));
    }
    let mut names = HashSet::new();
    for page in pages.iter_mut() {
        if page.name.is_empty() {
            return Err(invalid("style:name must not be empty"));
        }
        if !names.insert(page.name.clone()) {
            return Err(invalid("duplicate style:master-page name"));
        }
        page.children.clear();
    }

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<ExpandedName> = Vec::new();
    let mut active: Option<ActiveMaster> = None;
    let mut page_index = 0usize;
    let mut events = 0usize;
    loop {
        events += 1;
        if events > MAX_EVENTS {
            return Err(invalid("styles XML has too many events"));
        }
        let start = reader.buffer_position() as usize;
        let (_, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("styles XML parsing error: {error}")))?;
        let end = reader.buffer_position() as usize;
        match event.into_owned() {
            Event::Start(element) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid("styles XML exceeds 256 levels"));
                }
                let current = expanded(&reader, element.name());
                let parent_is_master_styles = stack
                    .last()
                    .is_some_and(|parent| is_name(parent, OFFICE, b"master-styles"));
                if is_name(&current, STYLE, b"master-page") {
                    if !parent_is_master_styles || active.is_some() {
                        return Err(invalid(
                            "style:master-page must be a direct office:master-styles child",
                        ));
                    }
                    let page = pages
                        .get(page_index)
                        .ok_or_else(|| invalid("master-page inventory mismatch"))?;
                    validate_header(&reader, &element, page)?;
                    active = Some(ActiveMaster {
                        page: page_index,
                        depth: stack.len() + 1,
                        last_rank: None,
                        seen: [false; 11],
                        child: None,
                        total_child_bytes: 0,
                    });
                    page_index += 1;
                } else if let Some(master) = active.as_mut() {
                    let depth = stack.len() + 1;
                    if depth == master.depth + 1 {
                        let kind = classify(&current)
                            .ok_or_else(|| invalid("unexpected direct style:master-page child"))?;
                        register_child(master, kind)?;
                        master.child = Some(ActiveChild { kind, start, depth });
                    }
                }
                stack.push(current);
            },
            Event::Empty(element) => {
                let current = expanded(&reader, element.name());
                let parent_is_master_styles = stack
                    .last()
                    .is_some_and(|parent| is_name(parent, OFFICE, b"master-styles"));
                if is_name(&current, STYLE, b"master-page") {
                    if !parent_is_master_styles || active.is_some() {
                        return Err(invalid(
                            "style:master-page must be a direct office:master-styles child",
                        ));
                    }
                    let page = pages
                        .get(page_index)
                        .ok_or_else(|| invalid("master-page inventory mismatch"))?;
                    validate_header(&reader, &element, page)?;
                    page_index += 1;
                } else if let Some(master) = active.as_mut()
                    && stack.len() + 1 == master.depth + 1
                {
                    let kind = classify(&current)
                        .ok_or_else(|| invalid("unexpected direct style:master-page child"))?;
                    register_child(master, kind)?;
                    push_child(xml, pages, master, kind, start, end)?;
                }
            },
            Event::End(_) => {
                let depth = stack.len();
                if let Some(master) = active.as_mut()
                    && master
                        .child
                        .as_ref()
                        .is_some_and(|child| child.depth == depth)
                {
                    let child = master.child.take().unwrap();
                    push_child(xml, pages, master, child.kind, child.start, end)?;
                }
                if active.as_ref().is_some_and(|master| master.depth == depth) {
                    if active.as_ref().unwrap().child.is_some() {
                        return Err(invalid("unterminated master-page child"));
                    }
                    active = None;
                }
                stack
                    .pop()
                    .ok_or_else(|| invalid("invalid styles XML nesting"))?;
            },
            Event::Text(value) => {
                let bytes: &[u8] = value.as_ref();
                if let Some(master) = &active
                    && stack.len() == master.depth
                    && !bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid(
                        "non-whitespace text is not allowed directly in style:master-page",
                    ));
                }
            },
            Event::CData(value) => {
                let bytes: &[u8] = value.as_ref();
                if let Some(master) = &active
                    && stack.len() == master.depth
                    && !bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid(
                        "non-whitespace CDATA is not allowed directly in style:master-page",
                    ));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are not allowed"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if active.is_some() || !stack.is_empty() || page_index != pages.len() {
        return Err(invalid("truncated master-page XML"));
    }
    Ok(())
}

fn canonical_fragment(page: &MasterPage) -> Result<String> {
    if page.name.is_empty() {
        return Err(invalid("master-page name must not be empty"));
    }
    let layout = page
        .page_layout_name
        .as_deref()
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid("master-page page-layout-name is required"))?;
    let escape = litchi_core::xml::escape_xml;
    let mut output = format!(
        "<style:master-page xmlns:style=\"{}\" xmlns:draw=\"{}\" xmlns:office=\"{}\" xmlns:presentation=\"{}\" xmlns:anim=\"{}\" xmlns:dr3d=\"{}\" style:name=\"{}\"",
        String::from_utf8_lossy(STYLE),
        String::from_utf8_lossy(DRAW),
        String::from_utf8_lossy(OFFICE),
        String::from_utf8_lossy(PRESENTATION),
        String::from_utf8_lossy(ANIM),
        String::from_utf8_lossy(DR3D),
        escape(&page.name),
    );
    if let Some(value) = &page.display_name {
        output.push_str(&format!(" style:display-name=\"{}\"", escape(value)));
    }
    output.push_str(&format!(" style:page-layout-name=\"{}\"", escape(layout)));
    if let Some(value) = &page.drawing_style_name {
        output.push_str(&format!(" draw:style-name=\"{}\"", escape(value)));
    }
    if let Some(value) = &page.next_style_name {
        output.push_str(&format!(" style:next-style-name=\"{}\"", escape(value)));
    }
    if page.children.is_empty() {
        output.push_str("/>");
        return Ok(output);
    }
    output.push('>');
    for child in &page.children {
        output.push_str(&child.xml);
    }
    output.push_str("</style:master-page>");
    let wrapper = format!(
        "<office:document xmlns:office=\"{}\"><office:master-styles>{}</office:master-styles></office:document>",
        String::from_utf8_lossy(OFFICE),
        output
    );
    let parsed = crate::odt::header_footer::parse_master_pages(&wrapper)?;
    if parsed.len() != 1 {
        return Err(invalid("canonical master-page did not validate"));
    }
    Ok(output)
}

impl MasterPage {
    /// Create an empty schema-valid master page.
    pub fn try_new(name: impl Into<String>, page_layout_name: impl Into<String>) -> Result<Self> {
        let value = Self {
            name: name.into(),
            display_name: None,
            page_layout_name: Some(page_layout_name.into()),
            drawing_style_name: None,
            next_style_name: None,
            regions: Vec::new(),
            children: Vec::new(),
            xml: String::new(),
        };
        canonical_fragment(&value)?;
        Ok(value)
    }

    /// Serialize known attributes and typed children in canonical RNG order.
    pub fn to_xml_fragment(&self) -> Result<String> {
        canonical_fragment(self)
    }
}

impl MasterPageChild {
    /// Create a typed inert child from an exact XML fragment.
    pub fn new(kind: MasterPageChildKind, xml: impl Into<String>) -> Self {
        Self {
            kind,
            xml: xml.into(),
        }
    }
}

fn validate_fragment(fragment: &str) -> Result<MasterPage> {
    if fragment.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("master-page fragment exceeds 16 MiB"));
    }
    let wrapper = format!(
        "<office:document xmlns:office=\"{}\"><office:master-styles>{fragment}</office:master-styles></office:document>",
        String::from_utf8_lossy(OFFICE)
    );
    let mut pages = crate::odt::header_footer::parse_master_pages(&wrapper)?;
    if pages.len() != 1 || pages[0].xml != fragment {
        return Err(invalid("fragment must be exactly one style:master-page"));
    }
    Ok(pages.remove(0))
}

/// Insert one exact master-page fragment under `office:master-styles`.
pub fn insert_master_page_xml(xml: &str, fragment: &str) -> Result<String> {
    let requested = validate_fragment(fragment)?;
    if crate::odt::header_footer::parse_master_pages(xml)?
        .iter()
        .any(|page| page.name == requested.name)
    {
        return Err(invalid(format!(
            "master page '{}' already exists",
            requested.name
        )));
    }
    crate::odt::header_footer::insert_container_child(xml, OFFICE, b"master-styles", fragment)
}

/// Replace one named master page with an exact validated fragment.
pub fn replace_master_page_xml(xml: &str, name: &str, fragment: &str) -> Result<String> {
    let requested = validate_fragment(fragment)?;
    if requested.name != name {
        return Err(invalid(
            "replacement master-page name does not match target",
        ));
    }
    crate::odt::header_footer::parse_master_pages(xml)?;
    let location = crate::odt::header_footer::find_master_page(xml, name)?
        .ok_or_else(|| invalid(format!("master page '{name}' does not exist")))?;
    Ok(crate::odt::header_footer::replace_range(
        xml,
        location.start,
        location.end,
        fragment,
    ))
}

/// Remove one named master page without rewriting surrounding XML.
pub fn remove_master_page_xml(xml: &str, name: &str) -> Result<String> {
    crate::odt::header_footer::parse_master_pages(xml)?;
    let location = crate::odt::header_footer::find_master_page(xml, name)?
        .ok_or_else(|| invalid(format!("master page '{name}' does not exist")))?;
    Ok(crate::odt::header_footer::replace_range(
        xml,
        location.start,
        location.end,
        "",
    ))
}

impl crate::OpenDocumentPackage {
    /// Parse inert master-page metadata from packaged `styles.xml`.
    pub fn master_pages(&self) -> Result<Vec<MasterPage>> {
        self.styles_xml()?.map_or_else(
            || Ok(Vec::new()),
            |xml| crate::odt::header_footer::parse_master_pages(&xml),
        )
    }
}

impl crate::FlatOpenDocument {
    /// Parse inert master-page metadata from a flat OpenDocument.
    pub fn master_pages(&self) -> Result<Vec<MasterPage>> {
        crate::odt::header_footer::parse_master_pages(self.xml())
    }
}
