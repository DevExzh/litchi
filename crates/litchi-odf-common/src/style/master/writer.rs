//! Bounded, lossless ODF master-page and header/footer writer.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

use super::reader::read;
use super::region::Kind;
use super::{Child, ChildKind, Master};

const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const OFFICE: &[u8] = OFFICE_NAMESPACE;
const STYLE: &[u8] = STYLE_NAMESPACE;
const DRAW: &[u8] = DRAW_NAMESPACE;
const ANIM: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:animation:1.0";
const DR3D: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
const PRESENTATION: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const MAX_FRAGMENT_BYTES: usize = 16 * 1_048_576;

struct ElementLocation {
    start: usize,
    end: usize,
    content_end: usize,
    qualified_name: String,
    empty: bool,
}

pub(crate) struct MasterLocation {
    pub(crate) start: usize,
    pub(crate) end: usize,
    content_start: usize,
    content_end: usize,
    qualified_name: String,
    empty: bool,
}

impl Master {
    /// Create an empty schema-valid master page.
    ///
    /// # Errors
    ///
    /// Returns an error when either required attribute is empty or invalid.
    pub fn new(name: impl Into<String>, page_layout_name: impl Into<String>) -> Result<Self> {
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
    ///
    /// # Errors
    ///
    /// Returns an error when the master page or one of its children is invalid.
    pub fn to_xml_fragment(&self) -> Result<String> {
        canonical_fragment(self)
    }
}

impl Child {
    /// Create a typed inert child from an exact XML fragment.
    pub fn new(kind: ChildKind, xml: impl Into<String>) -> Self {
        Self {
            kind,
            xml: xml.into(),
        }
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn byte_offset(offset: u64, context: &str) -> Result<usize> {
    usize::try_from(offset)
        .map_err(|error| invalid(format!("{context} byte offset is out of range: {error}")))
}

fn append_attribute(output: &mut String, name: &str, value: &str) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    output.push_str(&litchi_core::xml::escape_xml(value));
    output.push('"');
}

fn canonical_fragment(page: &Master) -> Result<String> {
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
        append_attribute(&mut output, "style:display-name", value);
    }
    append_attribute(&mut output, "style:page-layout-name", layout);
    if let Some(value) = &page.drawing_style_name {
        append_attribute(&mut output, "draw:style-name", value);
    }
    if let Some(value) = &page.next_style_name {
        append_attribute(&mut output, "style:next-style-name", value);
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
        "<office:document xmlns:office=\"{}\"><office:master-styles>{output}</office:master-styles></office:document>",
        String::from_utf8_lossy(OFFICE),
    );
    let parsed = read(&wrapper)?;
    if parsed.len() != 1 {
        return Err(invalid("canonical master-page did not validate"));
    }
    Ok(output)
}

fn validate_fragment(fragment: &str) -> Result<Master> {
    if fragment.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("master-page fragment exceeds 16 MiB"));
    }
    let wrapper = format!(
        "<office:document xmlns:office=\"{}\"><office:master-styles>{fragment}</office:master-styles></office:document>",
        String::from_utf8_lossy(OFFICE)
    );
    let mut pages = read(&wrapper)?;
    if pages.len() != 1 || pages[0].xml != fragment {
        return Err(invalid("fragment must be exactly one style:master-page"));
    }
    Ok(pages.remove(0))
}

/// Insert one exact master-page fragment under `office:master-styles`.
///
/// # Errors
///
/// Returns an error when either XML input is invalid, the fragment is not an
/// exact master page, or its name already exists.
pub fn insert(xml: &str, fragment: &str) -> Result<String> {
    let requested = validate_fragment(fragment)?;
    if read(xml)?.iter().any(|page| page.name == requested.name) {
        return Err(invalid(format!(
            "master page '{}' already exists",
            requested.name
        )));
    }
    insert_container_child(xml, OFFICE, b"master-styles", fragment)
}

/// Replace one named master page with an exact validated fragment.
///
/// # Errors
///
/// Returns an error when the XML or fragment is invalid, or the target master
/// page does not exist.
pub fn replace(xml: &str, name: &str, fragment: &str) -> Result<String> {
    let requested = validate_fragment(fragment)?;
    if requested.name != name {
        return Err(invalid(
            "replacement master-page name does not match target",
        ));
    }
    read(xml)?;
    let location = find_master_page(xml, name)?
        .ok_or_else(|| invalid(format!("master page '{name}' does not exist")))?;
    replace_range(xml, location.start, location.end, fragment)
}

/// Remove one named master page without rewriting surrounding XML.
///
/// # Errors
///
/// Returns an error when the XML is invalid or the target master page does not
/// exist.
pub fn remove(xml: &str, name: &str) -> Result<String> {
    read(xml)?;
    let location = find_master_page(xml, name)?
        .ok_or_else(|| invalid(format!("master page '{name}' does not exist")))?;
    replace_range(xml, location.start, location.end, "")
}

/// Sets or removes the text of one master-page header or footer region.
///
/// # Errors
///
/// Returns an error when the XML is invalid or the named master page cannot be
/// found.
pub fn set_text(
    xml: &str,
    master_page_name: &str,
    kind: Kind,
    text: Option<&str>,
) -> Result<String> {
    let replacement = text.map(|plain_text| {
        format!(
            "<style:{name} xmlns:style=\"{style}\" xmlns:text=\"{text_ns}\"><text:p>{value}</text:p></style:{name}>",
            name = kind.element_name(),
            style = String::from_utf8_lossy(STYLE_NAMESPACE),
            text_ns = String::from_utf8_lossy(TEXT_NAMESPACE),
            value = litchi_core::xml::escape_xml(plain_text),
        )
    });
    replace_region(xml, master_page_name, kind, replacement.as_deref())
}

/// Replaces one master-page header or footer with validated XML.
///
/// # Errors
///
/// Returns an error when the region XML or document XML is invalid, or the
/// target master page cannot be found.
pub fn set_xml(xml: &str, master_page_name: &str, kind: Kind, region_xml: &str) -> Result<String> {
    validate_region_xml(region_xml, kind)?;
    replace_region(xml, master_page_name, kind, Some(region_xml))
}

fn validate_region_xml(region_xml: &str, kind: Kind) -> Result<()> {
    let wrapper = format!(
        "<office:document-styles xmlns:office=\"{}\" xmlns:style=\"{}\"><office:master-styles><style:master-page style:name=\"validation\" style:page-layout-name=\"validation\">{region_xml}</style:master-page></office:master-styles></office:document-styles>",
        String::from_utf8_lossy(OFFICE_NAMESPACE),
        String::from_utf8_lossy(STYLE_NAMESPACE),
    );
    let pages = read(&wrapper)?;
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
    kind: Kind,
    replacement: Option<&str>,
) -> Result<String> {
    let pages = read(xml)?;
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
        let Some(replacement_xml) = replacement else {
            return Ok(xml.to_string());
        };
        let empty = &xml[location.start..location.end];
        let marker = empty.rfind("/>").ok_or_else(|| {
            Error::InvalidFormat("malformed empty style:master-page element".to_string())
        })?;
        let mut expanded = String::with_capacity(empty.len() + replacement_xml.len() + 32);
        expanded.push_str(&empty[..marker]);
        expanded.push('>');
        expanded.push_str(replacement_xml);
        expanded.push_str("</");
        expanded.push_str(&location.qualified_name);
        expanded.push('>');
        return replace_range(xml, location.start, location.end, &expanded);
    }

    if let Some(region) = page.region(kind) {
        let content = &xml[location.content_start..location.content_end];
        let relative = content.find(&region.xml).ok_or_else(|| {
            Error::InvalidFormat("header/footer XML is outside its master page".to_string())
        })?;
        let start = location.content_start + relative;
        let end = start + region.xml.len();
        return replace_range(xml, start, end, replacement.unwrap_or(""));
    }
    let Some(replacement_xml) = replacement else {
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
    replace_range(xml, insertion, insertion, replacement_xml)
}

/// Adds a master page and its page layout when needed.
///
/// # Errors
///
/// Returns an error when the document XML is invalid, a required name is empty,
/// or a master page with the requested name already exists.
pub fn add(xml: &str, name: &str, page_layout_name: &str) -> Result<String> {
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
    if read(xml)?.iter().any(|page| page.name == name) {
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
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
}

pub(crate) fn insert_container_child(
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
        return replace_range(xml, location.content_end, location.content_end, child);
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
    replace_range(xml, location.start, location.end, &expanded)
}

fn find_element(xml: &str, namespace: &[u8], local_name: &[u8]) -> Result<Option<ElementLocation>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut active: Option<(usize, usize, usize, String)> = None;
    loop {
        let event_start = byte_offset(reader.buffer_position(), "styles XML")?;
        let (resolved_namespace, parsed_event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("styles.xml parsing error: {error}")))?;
        let matches = bound_to(&resolved_namespace, namespace);
        let event = parsed_event.into_owned();
        let event_end = byte_offset(reader.buffer_position(), "styles XML")?;
        match event {
            Event::Start(element)
                if active.is_none() && matches && element.local_name().as_ref() == local_name =>
            {
                let qualified_name =
                    String::from_utf8(element.name().as_ref().to_vec()).map_err(|error| {
                        Error::InvalidFormat(format!("invalid element name: {error}"))
                    })?;
                active = Some((event_start, event_end, 1, qualified_name));
            },
            Event::Empty(element)
                if active.is_none() && matches && element.local_name().as_ref() == local_name =>
            {
                let qualified_name =
                    String::from_utf8(element.name().as_ref().to_vec()).map_err(|error| {
                        Error::InvalidFormat(format!("invalid element name: {error}"))
                    })?;
                return Ok(Some(ElementLocation {
                    start: event_start,
                    end: event_end,
                    content_end: event_end,
                    qualified_name,
                    empty: true,
                }));
            },
            Event::Start(_) if active.is_some() => {
                let current = active.as_mut().ok_or_else(|| {
                    Error::InvalidFormat("missing active styles.xml element".to_string())
                })?;
                current.2 = current.2.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("styles.xml element nesting overflows".to_string())
                })?;
            },
            Event::End(_) if active.is_some() => {
                let complete = {
                    let current = active.as_mut().ok_or_else(|| {
                        Error::InvalidFormat("missing active styles.xml element".to_string())
                    })?;
                    current.2 = current.2.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid element nesting".to_string())
                    })?;
                    current.2 == 0
                };
                if complete {
                    let (start, _, _, qualified_name) = active.take().ok_or_else(|| {
                        Error::InvalidFormat("missing completed styles.xml element".to_string())
                    })?;
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
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
}

pub(crate) fn find_master_page(xml: &str, expected_name: &str) -> Result<Option<MasterLocation>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut active: Option<(usize, usize, usize, String)> = None;
    loop {
        let event_start = byte_offset(reader.buffer_position(), "styles XML")?;
        let (namespace, parsed_event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("styles.xml parsing error: {error}")))?;
        let style_element = bound_to(&namespace, STYLE_NAMESPACE);
        let event = parsed_event.into_owned();
        let event_end = byte_offset(reader.buffer_position(), "styles XML")?;
        match event {
            Event::Start(element)
                if active.is_none()
                    && style_element
                    && element.local_name().as_ref() == b"master-page"
                    && style_attr(&reader, &element, b"name")?.as_deref()
                        == Some(expected_name) =>
            {
                let qualified_name =
                    String::from_utf8(element.name().as_ref().to_vec()).map_err(|error| {
                        Error::InvalidFormat(format!("invalid master-page name: {error}"))
                    })?;
                active = Some((event_start, event_end, 1, qualified_name));
            },
            Event::Empty(element)
                if active.is_none()
                    && style_element
                    && element.local_name().as_ref() == b"master-page"
                    && style_attr(&reader, &element, b"name")?.as_deref()
                        == Some(expected_name) =>
            {
                let qualified_name =
                    String::from_utf8(element.name().as_ref().to_vec()).map_err(|error| {
                        Error::InvalidFormat(format!("invalid master-page name: {error}"))
                    })?;
                return Ok(Some(MasterLocation {
                    start: event_start,
                    end: event_end,
                    content_start: event_end,
                    content_end: event_end,
                    qualified_name,
                    empty: true,
                }));
            },
            Event::Start(_) if active.is_some() => {
                let current = active.as_mut().ok_or_else(|| {
                    Error::InvalidFormat("missing active style:master-page".to_string())
                })?;
                current.2 = current.2.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("master-page nesting overflows".to_string())
                })?;
            },
            Event::End(_) if active.is_some() => {
                let complete = {
                    let current = active.as_mut().ok_or_else(|| {
                        Error::InvalidFormat("missing active style:master-page".to_string())
                    })?;
                    current.2 = current.2.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid master-page nesting".to_string())
                    })?;
                    current.2 == 0
                };
                if complete {
                    let (start, content_start, _, qualified_name) =
                        active.take().ok_or_else(|| {
                            Error::InvalidFormat("missing completed style:master-page".to_string())
                        })?;
                    return Ok(Some(MasterLocation {
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
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
}

/// Replace one checked UTF-8 range without rebuilding unrelated XML.
///
/// # Errors
///
/// Returns an error when the requested range is outside the XML or splits a
/// UTF-8 character.
pub fn replace_range(xml: &str, start: usize, end: usize, replacement: &str) -> Result<String> {
    if start > end || end > xml.len() || !xml.is_char_boundary(start) || !xml.is_char_boundary(end)
    {
        return Err(invalid("XML splice range is not a UTF-8 boundary"));
    }
    let mut output = String::with_capacity(xml.len() - (end - start) + replacement.len());
    output.push_str(&xml[..start]);
    output.push_str(replacement);
    output.push_str(&xml[end..]);
    Ok(output)
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
    for raw_attribute in element.attributes() {
        let attribute = raw_attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid style attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == expected_namespace)
            && local.as_ref() == local_name
        {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|decoded_value| Some(decoded_value.into_owned()))
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid style attribute: {error}"))
                });
        }
    }
    Ok(None)
}

fn bound_to(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

#[cfg(test)]
mod tests {
    use super::*;

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    fn test_ok<T>(result: Result<T>) -> T {
        match result {
            Ok(value) => value,
            Err(error) => panic!("test operation failed: {error}"),
        }
    }

    fn test_some<T>(value: Option<T>) -> T {
        match value {
            Some(found_value) => found_value,
            None => panic!("test fixture did not contain a required value"),
        }
    }

    #[test]
    fn authors_and_round_trips_a_master() {
        let mut master = test_ok(Master::new("A", "pm1"));
        master.children.push(Child::new(
            ChildKind::Shape,
            r#"<d:rect xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"/>"#.to_string(),
        ));
        let fragment = test_ok(master.to_xml_fragment());
        assert!(fragment.contains(r#"style:name="A""#));
        assert!(read(&format!(
            r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:s="{STYLE}"><o:master-styles>{fragment}</o:master-styles></o:document-styles>"#
        )).is_ok());
    }

    #[test]
    fn updates_only_the_requested_region() {
        let styles = format!(
            r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:s="{STYLE}" xmlns:t="{TEXT}"><o:styles><s:style s:name="keep"/></o:styles><o:master-styles><s:master-page s:name="A" s:page-layout-name="pm1"><s:header><t:p>Old</t:p></s:header></s:master-page></o:master-styles></o:document-styles>"#
        );
        let updated = test_ok(set_text(&styles, "A", Kind::Header, Some("New & <value>")));
        assert!(updated.contains(r#"<s:style s:name="keep"/>"#));
        let pages = test_ok(read(&updated));
        assert_eq!(
            test_some(pages[0].region(Kind::Header)).text,
            "New & <value>"
        );
        let cleared = test_ok(set_text(&updated, "A", Kind::Header, None));
        assert!(test_ok(read(&cleared))[0].region(Kind::Header).is_none());
    }
}
