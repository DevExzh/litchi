//! Master-page headers and footers from ODT `styles.xml`.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesRef, BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

use super::header_footer_content::{
    HeaderFooterBlock, MAX_EXPANDED_SPACES, parse_header_footer_blocks,
};

const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";

/// One of the six header/footer regions supported by an ODF master page.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum HeaderFooterKind {
    Header,
    HeaderFirst,
    HeaderLeft,
    Footer,
    FooterFirst,
    FooterLeft,
}

impl HeaderFooterKind {
    pub(super) fn parse(local_name: &[u8]) -> Option<Self> {
        match local_name {
            b"header" => Some(Self::Header),
            b"header-first" => Some(Self::HeaderFirst),
            b"header-left" => Some(Self::HeaderLeft),
            b"footer" => Some(Self::Footer),
            b"footer-first" => Some(Self::FooterFirst),
            b"footer-left" => Some(Self::FooterLeft),
            _ => None,
        }
    }

    fn element_name(self) -> &'static str {
        match self {
            Self::Header => "header",
            Self::HeaderFirst => "header-first",
            Self::HeaderLeft => "header-left",
            Self::Footer => "footer",
            Self::FooterFirst => "footer-first",
            Self::FooterLeft => "footer-left",
        }
    }

    fn order(self) -> u8 {
        match self {
            Self::Header => 0,
            Self::HeaderLeft => 1,
            Self::HeaderFirst => 2,
            Self::Footer => 3,
            Self::FooterLeft => 4,
            Self::FooterFirst => 5,
        }
    }
}

/// Losslessly retained content of one master-page header or footer.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HeaderFooter {
    pub kind: HeaderFooterKind,
    /// The exact element bytes from `styles.xml`, including nested fields and formatting.
    pub xml: String,
    /// Best-effort visible literal text. Dynamic field values remain represented in `xml`.
    pub text: String,
    /// Ordered paragraphs/headings with explicit inline text, whitespace, and fields.
    pub blocks: Vec<HeaderFooterBlock>,
}

/// An ODF master page and all of its header/footer regions.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MasterPage {
    pub name: String,
    pub display_name: Option<String>,
    pub page_layout_name: Option<String>,
    pub drawing_style_name: Option<String>,
    pub next_style_name: Option<String>,
    pub regions: Vec<HeaderFooter>,
    /// The exact master-page element bytes, including shapes and extension content.
    pub xml: String,
}

impl MasterPage {
    /// Return a particular header/footer region when it exists.
    pub fn region(&self, kind: HeaderFooterKind) -> Option<&HeaderFooter> {
        self.regions.iter().find(|region| region.kind == kind)
    }
}

struct MasterPageBuilder {
    page: MasterPage,
    start: usize,
    depth: usize,
}

struct RegionBuilder {
    kind: HeaderFooterKind,
    start: usize,
    depth: usize,
    text: String,
    expanded_spaces: usize,
}

pub(crate) fn parse_master_pages(xml: &str) -> Result<Vec<MasterPage>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut pages = Vec::new();
    let mut master: Option<MasterPageBuilder> = None;
    let mut region: Option<RegionBuilder> = None;

    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("styles.xml parsing error: {error}")))?;
        let style_element = bound_to(&namespace, STYLE_NAMESPACE);
        let text_element = bound_to(&namespace, TEXT_NAMESPACE);
        let event = event.into_owned();
        let event_end = reader.buffer_position() as usize;

        match event {
            Event::Start(element)
                if style_element && element.local_name().as_ref() == b"master-page" =>
            {
                if master.is_some() {
                    return Err(Error::InvalidFormat(
                        "nested style:master-page element".to_string(),
                    ));
                }
                master = Some(MasterPageBuilder {
                    page: parse_master_page(&reader, &element)?,
                    start: event_start,
                    depth: 1,
                });
            },
            Event::Empty(element)
                if style_element && element.local_name().as_ref() == b"master-page" =>
            {
                let mut page = parse_master_page(&reader, &element)?;
                page.xml = xml[event_start..event_end].to_string();
                pages.push(page);
            },
            Event::Start(element) if master.is_some() => {
                let master = master.as_mut().expect("checked master page");
                master.depth += 1;
                if region.is_none()
                    && style_element
                    && let Some(kind) = HeaderFooterKind::parse(element.local_name().as_ref())
                {
                    if master.page.region(kind).is_some() {
                        return Err(Error::InvalidFormat(format!(
                            "duplicate {kind:?} in master page '{}'",
                            master.page.name
                        )));
                    }
                    region = Some(RegionBuilder {
                        kind,
                        start: event_start,
                        depth: 1,
                        text: String::new(),
                        expanded_spaces: 0,
                    });
                } else if let Some(region) = region.as_mut() {
                    region.depth += 1;
                }
            },
            Event::Empty(element)
                if master.is_some()
                    && region.is_none()
                    && style_element
                    && HeaderFooterKind::parse(element.local_name().as_ref()).is_some() =>
            {
                let kind = HeaderFooterKind::parse(element.local_name().as_ref()).unwrap();
                let master = master.as_mut().expect("checked master page");
                if master.page.region(kind).is_some() {
                    return Err(Error::InvalidFormat(format!(
                        "duplicate {kind:?} in master page '{}'",
                        master.page.name
                    )));
                }
                master.page.regions.push(HeaderFooter {
                    kind,
                    xml: xml[event_start..event_end].to_string(),
                    text: String::new(),
                    blocks: Vec::new(),
                });
            },
            Event::Empty(element) if region.is_some() && text_element => {
                let region = region.as_mut().expect("checked region");
                append_empty_text_element(
                    &reader,
                    &element,
                    &mut region.text,
                    &mut region.expanded_spaces,
                )?;
            },
            Event::Text(value) if region.is_some() => {
                let decoded = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid header text: {error}"))
                    })?;
                region.as_mut().unwrap().text.push_str(&decoded);
            },
            Event::GeneralRef(reference) if region.is_some() => {
                region
                    .as_mut()
                    .unwrap()
                    .text
                    .push_str(&decode_reference(&reference)?);
            },
            Event::CData(value) if region.is_some() => {
                let decoded = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid header CDATA: {error}"))
                    })?;
                region.as_mut().unwrap().text.push_str(&decoded);
            },
            Event::End(element) if master.is_some() => {
                if let Some(active) = region.as_mut() {
                    if text_element && matches!(element.local_name().as_ref(), b"p" | b"h") {
                        active.text.push('\n');
                    }
                    active.depth = active.depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid header/footer nesting".to_string())
                    })?;
                    if active.depth == 0 {
                        let active = region.take().expect("checked region");
                        let master = master.as_mut().expect("checked master page");
                        master.page.regions.push(HeaderFooter {
                            kind: active.kind,
                            xml: xml[active.start..event_end].to_string(),
                            text: active.text.trim_end_matches('\n').to_string(),
                            blocks: Vec::new(),
                        });
                    }
                }
                let current = master.as_mut().expect("checked master page");
                current.depth = current.depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid master-page nesting".to_string())
                })?;
                if current.depth == 0 {
                    if !style_element || element.local_name().as_ref() != b"master-page" {
                        return Err(Error::InvalidFormat(
                            "malformed style:master-page element".to_string(),
                        ));
                    }
                    let mut finished = master.take().expect("checked master page");
                    finished.page.xml = xml[finished.start..event_end].to_string();
                    pages.push(finished.page);
                }
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DTD is not allowed in ODF styles.xml".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if master.is_some() || region.is_some() {
        return Err(Error::InvalidFormat(
            "unterminated master-page header/footer".to_string(),
        ));
    }
    let mut structured = parse_header_footer_blocks(xml)?;
    for page in &mut pages {
        for region in &mut page.regions {
            region.blocks = structured
                .remove(&(page.name.clone(), region.kind))
                .unwrap_or_default();
        }
    }
    Ok(pages)
}

pub(crate) fn set_region_text(
    xml: &str,
    master_page_name: &str,
    kind: HeaderFooterKind,
    text: Option<&str>,
) -> Result<String> {
    let replacement = text.map(|text| {
        format!(
            "<style:{name} xmlns:style=\"{style}\" xmlns:text=\"{text_ns}\"><text:p>{value}</text:p></style:{name}>",
            name = kind.element_name(),
            style = String::from_utf8_lossy(STYLE_NAMESPACE),
            text_ns = String::from_utf8_lossy(TEXT_NAMESPACE),
            value = litchi_core::xml::escape_xml(text),
        )
    });
    replace_region(xml, master_page_name, kind, replacement.as_deref())
}

pub(crate) fn set_region_xml(
    xml: &str,
    master_page_name: &str,
    kind: HeaderFooterKind,
    region_xml: &str,
) -> Result<String> {
    validate_region_xml(region_xml, kind)?;
    replace_region(xml, master_page_name, kind, Some(region_xml))
}

fn validate_region_xml(region_xml: &str, kind: HeaderFooterKind) -> Result<()> {
    let wrapper = format!(
        "<office:document-styles xmlns:office=\"{}\" xmlns:style=\"{}\"><office:master-styles><style:master-page style:name=\"validation\">{region_xml}</style:master-page></office:master-styles></office:document-styles>",
        String::from_utf8_lossy(OFFICE_NAMESPACE),
        String::from_utf8_lossy(STYLE_NAMESPACE),
    );
    let pages = parse_master_pages(&wrapper)?;
    let valid = pages.len() == 1
        && pages[0].regions.len() == 1
        && pages[0].regions[0].kind == kind
        && pages[0].regions[0].xml == region_xml;
    if !valid {
        return Err(Error::InvalidFormat(format!(
            "header/footer XML must be exactly one style:{} element",
            kind.element_name()
        )));
    }
    Ok(())
}

fn replace_region(
    xml: &str,
    master_page_name: &str,
    kind: HeaderFooterKind,
    replacement: Option<&str>,
) -> Result<String> {
    let pages = parse_master_pages(xml)?;
    let page = pages
        .iter()
        .find(|page| page.name == master_page_name)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("master page '{master_page_name}' does not exist"))
        })?;
    let location = find_master_page(xml, master_page_name)?.ok_or_else(|| {
        Error::InvalidFormat(format!("master page '{master_page_name}' does not exist"))
    })?;

    if location.empty {
        let Some(replacement) = replacement else {
            return Ok(xml.to_string());
        };
        let empty = &xml[location.start..location.end];
        let marker = empty.rfind("/>").ok_or_else(|| {
            Error::InvalidFormat("malformed empty style:master-page element".to_string())
        })?;
        let mut expanded = String::with_capacity(empty.len() + replacement.len() + 32);
        expanded.push_str(&empty[..marker]);
        expanded.push('>');
        expanded.push_str(replacement);
        expanded.push_str("</");
        expanded.push_str(&location.qualified_name);
        expanded.push('>');
        return Ok(replace_range(xml, location.start, location.end, &expanded));
    }

    if let Some(region) = page.region(kind) {
        let content = &xml[location.content_start..location.content_end];
        let relative = content.find(&region.xml).ok_or_else(|| {
            Error::InvalidFormat("header/footer XML is outside its master page".to_string())
        })?;
        let start = location.content_start + relative;
        let end = start + region.xml.len();
        return Ok(replace_range(xml, start, end, replacement.unwrap_or("")));
    }
    let Some(replacement) = replacement else {
        return Ok(xml.to_string());
    };
    let mut insertion = location.content_start;
    for existing in &page.regions {
        let content = &xml[location.content_start..location.content_end];
        let relative = content.find(&existing.xml).ok_or_else(|| {
            Error::InvalidFormat("header/footer XML is outside its master page".to_string())
        })?;
        let start = location.content_start + relative;
        if existing.kind.order() > kind.order() {
            insertion = start;
            break;
        }
        insertion = start + existing.xml.len();
    }
    Ok(replace_range(xml, insertion, insertion, replacement))
}

pub(crate) fn add_master_page(xml: &str, name: &str, page_layout_name: &str) -> Result<String> {
    if name.is_empty() {
        return Err(Error::InvalidFormat(
            "master-page name must not be empty".to_string(),
        ));
    }
    if page_layout_name.is_empty() {
        return Err(Error::InvalidFormat(
            "page-layout name must not be empty".to_string(),
        ));
    }
    if parse_master_pages(xml)?
        .iter()
        .any(|page| page.name == name)
    {
        return Err(Error::InvalidFormat(format!(
            "master page '{name}' already exists"
        )));
    }

    let mut output = xml.to_string();
    if !has_named_style_element(&output, b"page-layout", page_layout_name)? {
        let layout = format!(
            "<style:page-layout xmlns:style=\"{}\" style:name=\"{}\"/>",
            String::from_utf8_lossy(STYLE_NAMESPACE),
            litchi_core::xml::escape_xml(page_layout_name),
        );
        output = insert_container_child(&output, OFFICE_NAMESPACE, b"automatic-styles", &layout)?;
    }
    let master = format!(
        "<style:master-page xmlns:style=\"{}\" style:name=\"{}\" style:page-layout-name=\"{}\"/>",
        String::from_utf8_lossy(STYLE_NAMESPACE),
        litchi_core::xml::escape_xml(name),
        litchi_core::xml::escape_xml(page_layout_name),
    );
    insert_container_child(&output, OFFICE_NAMESPACE, b"master-styles", &master)
}

fn has_named_style_element(xml: &str, local_name: &[u8], expected_name: &str) -> Result<bool> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("styles.xml parsing error: {error}")))?;
        let style_element = bound_to(&namespace, STYLE_NAMESPACE);
        match event {
            Event::Start(element) | Event::Empty(element)
                if style_element
                    && element.local_name().as_ref() == local_name
                    && style_attr(&reader, &element, b"name")?.as_deref()
                        == Some(expected_name) =>
            {
                return Ok(true);
            },
            Event::Eof => return Ok(false),
            _ => {},
        }
        buffer.clear();
    }
}

struct ElementLocation {
    start: usize,
    end: usize,
    content_end: usize,
    qualified_name: String,
    empty: bool,
}

fn insert_container_child(
    xml: &str,
    namespace: &[u8],
    local_name: &[u8],
    child: &str,
) -> Result<String> {
    let location = find_element(xml, namespace, local_name)?.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "styles.xml is missing {}",
            String::from_utf8_lossy(local_name)
        ))
    })?;
    if !location.empty {
        return Ok(replace_range(
            xml,
            location.content_end,
            location.content_end,
            child,
        ));
    }
    let empty = &xml[location.start..location.end];
    let marker = empty.rfind("/>").ok_or_else(|| {
        Error::InvalidFormat(format!(
            "malformed empty {} element",
            String::from_utf8_lossy(local_name)
        ))
    })?;
    let mut expanded = String::with_capacity(empty.len() + child.len() + 32);
    expanded.push_str(&empty[..marker]);
    expanded.push('>');
    expanded.push_str(child);
    expanded.push_str("</");
    expanded.push_str(&location.qualified_name);
    expanded.push('>');
    Ok(replace_range(xml, location.start, location.end, &expanded))
}

fn find_element(xml: &str, namespace: &[u8], local_name: &[u8]) -> Result<Option<ElementLocation>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut active: Option<(usize, usize, usize, String)> = None;
    loop {
        let event_start = reader.buffer_position() as usize;
        let (resolved_namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("styles.xml parsing error: {error}")))?;
        let matches = bound_to(&resolved_namespace, namespace);
        let event = event.into_owned();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(element)
                if active.is_none() && matches && element.local_name().as_ref() == local_name =>
            {
                let qualified_name = String::from_utf8(element.name().as_ref().to_vec())
                    .map_err(|_| Error::InvalidFormat("invalid element name".to_string()))?;
                active = Some((event_start, event_end, 1, qualified_name));
            },
            Event::Empty(element)
                if active.is_none() && matches && element.local_name().as_ref() == local_name =>
            {
                let qualified_name = String::from_utf8(element.name().as_ref().to_vec())
                    .map_err(|_| Error::InvalidFormat("invalid element name".to_string()))?;
                return Ok(Some(ElementLocation {
                    start: event_start,
                    end: event_end,
                    content_end: event_end,
                    qualified_name,
                    empty: true,
                }));
            },
            Event::Start(_) if active.is_some() => active.as_mut().unwrap().2 += 1,
            Event::End(_) if active.is_some() => {
                let current = active.as_mut().unwrap();
                current.2 = current
                    .2
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid element nesting".to_string()))?;
                if current.2 == 0 {
                    let (start, _, _, qualified_name) = active.take().unwrap();
                    return Ok(Some(ElementLocation {
                        start,
                        end: event_end,
                        content_end: event_start,
                        qualified_name,
                        empty: false,
                    }));
                }
            },
            Event::Eof => return Ok(None),
            _ => {},
        }
        buffer.clear();
    }
}

struct MasterPageLocation {
    start: usize,
    end: usize,
    content_start: usize,
    content_end: usize,
    qualified_name: String,
    empty: bool,
}

fn find_master_page(xml: &str, expected_name: &str) -> Result<Option<MasterPageLocation>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut active: Option<(usize, usize, usize, String)> = None;
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("styles.xml parsing error: {error}")))?;
        let style_element = bound_to(&namespace, STYLE_NAMESPACE);
        let event = event.into_owned();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(element)
                if active.is_none()
                    && style_element
                    && element.local_name().as_ref() == b"master-page"
                    && style_attr(&reader, &element, b"name")?.as_deref()
                        == Some(expected_name) =>
            {
                let qualified_name = String::from_utf8(element.name().as_ref().to_vec())
                    .map_err(|_| Error::InvalidFormat("invalid master-page name".to_string()))?;
                active = Some((event_start, event_end, 1, qualified_name));
            },
            Event::Empty(element)
                if active.is_none()
                    && style_element
                    && element.local_name().as_ref() == b"master-page"
                    && style_attr(&reader, &element, b"name")?.as_deref()
                        == Some(expected_name) =>
            {
                let qualified_name = String::from_utf8(element.name().as_ref().to_vec())
                    .map_err(|_| Error::InvalidFormat("invalid master-page name".to_string()))?;
                return Ok(Some(MasterPageLocation {
                    start: event_start,
                    end: event_end,
                    content_start: event_end,
                    content_end: event_end,
                    qualified_name,
                    empty: true,
                }));
            },
            Event::Start(_) if active.is_some() => active.as_mut().unwrap().2 += 1,
            Event::End(_) if active.is_some() => {
                let current = active.as_mut().unwrap();
                current.2 = current.2.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid master-page nesting".to_string())
                })?;
                if current.2 == 0 {
                    let (start, content_start, _, qualified_name) = active.take().unwrap();
                    return Ok(Some(MasterPageLocation {
                        start,
                        end: event_end,
                        content_start,
                        content_end: event_start,
                        qualified_name,
                        empty: false,
                    }));
                }
            },
            Event::Eof => return Ok(None),
            _ => {},
        }
        buffer.clear();
    }
}

fn replace_range(xml: &str, start: usize, end: usize, replacement: &str) -> String {
    let mut output = String::with_capacity(xml.len() - (end - start) + replacement.len());
    output.push_str(&xml[..start]);
    output.push_str(replacement);
    output.push_str(&xml[end..]);
    output
}

fn parse_master_page(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<MasterPage> {
    let name = style_attr(reader, element, b"name")?.ok_or_else(|| {
        Error::InvalidFormat("style:master-page is missing style:name".to_string())
    })?;
    Ok(MasterPage {
        name,
        display_name: style_attr(reader, element, b"display-name")?,
        page_layout_name: style_attr(reader, element, b"page-layout-name")?,
        drawing_style_name: namespaced_attr(reader, element, DRAW_NAMESPACE, b"style-name")?,
        next_style_name: style_attr(reader, element, b"next-style-name")?,
        regions: Vec::new(),
        xml: String::new(),
    })
}

fn style_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<Option<String>> {
    namespaced_attr(reader, element, STYLE_NAMESPACE, local_name)
}

fn namespaced_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    local_name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid master-page attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == expected_namespace)
            && local.as_ref() == local_name
        {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid style attribute: {error}"))
                });
        }
    }
    Ok(None)
}

fn append_empty_text_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    output: &mut String,
    expanded_spaces: &mut usize,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"s" => {
            let count = style_independent_text_count(reader, element)?.unwrap_or(1);
            *expanded_spaces = expanded_spaces.checked_add(count).ok_or_else(|| {
                Error::InvalidFormat("header text:s count exceeds safety limit".to_string())
            })?;
            if *expanded_spaces > MAX_EXPANDED_SPACES {
                return Err(Error::InvalidFormat(
                    "header text:s count exceeds safety limit".to_string(),
                ));
            }
            output.extend(std::iter::repeat_n(' ', count));
        },
        b"tab" => output.push('\t'),
        b"line-break" => output.push('\n'),
        _ => {},
    }
    Ok(())
}

fn style_independent_text_count(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Option<usize>> {
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid text:s attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == TEXT_NAMESPACE)
            && local.as_ref() == b"c"
        {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())?;
            return value.parse().map(Some).map_err(|_| {
                Error::InvalidFormat("invalid text:c count in header/footer".to_string())
            });
        }
    }
    Ok(None)
}

fn bound_to(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid header character reference: {error}"))
    })? {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid header entity: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "quot" => Ok("\"".to_string()),
        "apos" => Ok("'".to_string()),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported header entity '&{name};'"
        ))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_all_master_page_regions_losslessly_with_arbitrary_prefixes() {
        let xml = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:master-styles><s:master-page s:name="Standard" s:display-name="Default &amp; Main" s:page-layout-name="pm1" d:style-name="drawing1" s:next-style-name="Next"><s:header><t:p>Page <t:page-number/></t:p><t:p>A<t:s t:c="2"/>B<t:tab/>C<t:line-break/>D</t:p></s:header><s:header-first><t:p>First</t:p></s:header-first><s:header-left><t:p>Left</t:p></s:header-left><s:footer><t:p>Footer</t:p></s:footer><s:footer-first><t:p>First footer</t:p></s:footer-first><s:footer-left><t:p>Left footer</t:p></s:footer-left></s:master-page><s:master-page s:name="Empty"/></o:master-styles></o:document-styles>"#;
        let pages = parse_master_pages(xml).unwrap();
        assert_eq!(pages.len(), 2);
        assert_eq!(pages[0].name, "Standard");
        assert_eq!(pages[0].display_name.as_deref(), Some("Default & Main"));
        assert_eq!(pages[0].page_layout_name.as_deref(), Some("pm1"));
        assert_eq!(pages[0].drawing_style_name.as_deref(), Some("drawing1"));
        assert_eq!(pages[0].next_style_name.as_deref(), Some("Next"));
        assert_eq!(pages[0].regions.len(), 6);
        let header = pages[0].region(HeaderFooterKind::Header).unwrap();
        assert_eq!(header.text, "Page \nA  B\tC\nD");
        assert!(header.xml.starts_with("<s:header>"));
        assert!(header.xml.contains("<t:page-number/>"));
        assert!(pages[0].xml.starts_with("<s:master-page"));
        assert!(pages[0].xml.ends_with("</s:master-page>"));
        assert_eq!(pages[1].name, "Empty");
        assert!(pages[1].regions.is_empty());
        assert_eq!(pages[1].xml, "<s:master-page s:name=\"Empty\"/>");
    }

    #[test]
    fn rejects_duplicate_regions_and_missing_master_names() {
        let duplicate = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><o:master-styles><s:master-page s:name="A"><s:header/><s:header/></s:master-page></o:master-styles></o:document-styles>"#;
        assert!(parse_master_pages(duplicate).is_err());
        let missing = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><o:master-styles><s:master-page/></o:master-styles></o:document-styles>"#;
        assert!(parse_master_pages(missing).is_err());
    }

    #[test]
    fn inserts_replaces_and_clears_regions_without_rewriting_other_styles() {
        let xml = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:styles><s:style s:name="keep"/></o:styles><o:master-styles><s:master-page s:name="A"><s:header><t:p>Old</t:p></s:header><s:region-left/></s:master-page><s:master-page s:name="B" /></o:master-styles></o:document-styles>"#;
        let replaced =
            set_region_text(xml, "A", HeaderFooterKind::Header, Some("A & <B>")).unwrap();
        assert!(replaced.contains("<s:style s:name=\"keep\"/>"));
        assert!(replaced.contains("<s:region-left/>"));
        let pages = parse_master_pages(&replaced).unwrap();
        assert_eq!(
            pages[0].region(HeaderFooterKind::Header).unwrap().text,
            "A & <B>"
        );

        let inserted =
            set_region_text(&replaced, "A", HeaderFooterKind::FooterLeft, Some("Left")).unwrap();
        let pages = parse_master_pages(&inserted).unwrap();
        assert_eq!(
            pages[0].region(HeaderFooterKind::FooterLeft).unwrap().text,
            "Left"
        );

        let expanded =
            set_region_text(&inserted, "B", HeaderFooterKind::Footer, Some("B footer")).unwrap();
        assert!(expanded.contains("</s:master-page>"));
        let pages = parse_master_pages(&expanded).unwrap();
        assert_eq!(
            pages[1].region(HeaderFooterKind::Footer).unwrap().text,
            "B footer"
        );

        let cleared = set_region_text(&expanded, "A", HeaderFooterKind::Header, None).unwrap();
        let pages = parse_master_pages(&cleared).unwrap();
        assert!(pages[0].region(HeaderFooterKind::Header).is_none());
        assert!(cleared.contains("<s:region-left/>"));
    }

    #[test]
    fn adds_master_pages_and_reuses_page_layouts_with_arbitrary_prefixes() {
        let xml = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><o:automatic-styles /><o:master-styles /></o:document-styles>"#;
        let first = add_master_page(xml, "First & Main", "pm&1").unwrap();
        let second = add_master_page(&first, "Second", "pm&1").unwrap();
        let pages = parse_master_pages(&second).unwrap();

        assert_eq!(pages.len(), 2);
        assert_eq!(pages[0].name, "First & Main");
        assert_eq!(pages[0].page_layout_name.as_deref(), Some("pm&1"));
        assert_eq!(pages[1].name, "Second");
        assert_eq!(second.matches("<style:page-layout ").count(), 1);
        assert!(second.contains("</o:automatic-styles>"));
        assert!(second.contains("</o:master-styles>"));
        assert!(add_master_page(&second, "Second", "pm2").is_err());
        assert!(add_master_page(&second, "", "pm2").is_err());
        assert!(add_master_page(&second, "Third", "").is_err());
    }

    #[test]
    fn validates_and_sets_complete_rich_region_fragments() {
        let styles = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><o:master-styles><s:master-page s:name="A" s:page-layout-name="pm1"/></o:master-styles></o:document-styles>"#;
        let rich = r#"<x:header xmlns:x="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:p>Page <t:page-number/> of <t:page-count/></t:p></x:header>"#;
        let updated = set_region_xml(styles, "A", HeaderFooterKind::Header, rich).unwrap();
        let pages = parse_master_pages(&updated).unwrap();
        let header = pages[0].region(HeaderFooterKind::Header).unwrap();
        assert_eq!(header.xml, rich);
        assert_eq!(header.text, "Page  of ");

        assert!(set_region_xml(styles, "A", HeaderFooterKind::Footer, rich).is_err());
        assert!(set_region_xml(styles, "A", HeaderFooterKind::Header, " malformed ").is_err());
        assert!(
            set_region_xml(
                styles,
                "A",
                HeaderFooterKind::Header,
                &format!("{rich}{rich}"),
            )
            .is_err()
        );
    }

    #[test]
    fn opens_libreoffice_first_left_right_header_footer_fixture() {
        let document = crate::Document::open(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../3rdparty/libreoffice-core/sw/qa/core/header_footer/data/first-header-footer.odt"
        ))
        .unwrap();
        let pages = document.master_pages().unwrap();
        assert_eq!(pages.len(), 2);
        for kind in [
            HeaderFooterKind::Header,
            HeaderFooterKind::HeaderFirst,
            HeaderFooterKind::HeaderLeft,
            HeaderFooterKind::Footer,
            HeaderFooterKind::FooterFirst,
            HeaderFooterKind::FooterLeft,
        ] {
            let matching: Vec<_> = pages
                .iter()
                .filter_map(|page| page.region(kind))
                .collect();
            assert_eq!(matching.len(), 2, "missing {kind:?} regions");
            assert!(matching.iter().all(|region| !region.text.is_empty()));
            assert!(matching.iter().all(|region| region.blocks.len() == 1));
        }
    }
}
