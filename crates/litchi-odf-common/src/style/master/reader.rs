//! Bounded ODF master-page and header/footer reader.

use std::collections::HashSet;

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesRef, BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

use super::content::{MAX_EXPANDED_SPACES, parse};
use super::region::{Kind, Region};
use super::{Child, ChildKind, Master};

const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const OFFICE: &[u8] = OFFICE_NAMESPACE;
const STYLE: &[u8] = STYLE_NAMESPACE;
const DRAW: &[u8] = DRAW_NAMESPACE;
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
    kind: ChildKind,
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

fn classify(value: &ExpandedName) -> Option<ChildKind> {
    if value.namespace.as_deref() == Some(STYLE) {
        return Some(match value.local.as_slice() {
            b"header" => ChildKind::Region(Kind::Header),
            b"header-left" => ChildKind::Region(Kind::HeaderLeft),
            b"header-first" => ChildKind::Region(Kind::HeaderFirst),
            b"footer" => ChildKind::Region(Kind::Footer),
            b"footer-left" => ChildKind::Region(Kind::FooterLeft),
            b"footer-first" => ChildKind::Region(Kind::FooterFirst),
            _ => return None,
        });
    }
    if is_name(value, DRAW, b"layer-set") {
        return Some(ChildKind::LayerSet);
    }
    if is_name(value, OFFICE, b"forms") {
        return Some(ChildKind::Forms);
    }
    if is_name(value, PRESENTATION, b"notes") {
        return Some(ChildKind::Notes);
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
        return Some(ChildKind::Animation);
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
        return Some(ChildKind::Shape);
    }
    None
}

fn register_child(active: &mut ActiveMaster, kind: ChildKind) -> Result<()> {
    if active.child.is_some() {
        return Err(invalid("master-page direct children overlap"));
    }
    let rank = kind.order();
    if active.last_rank.is_some_and(|previous| rank < previous) {
        return Err(invalid(
            "style:master-page children are out of ODF 1.3 schema order",
        ));
    }
    if kind != ChildKind::Shape && active.seen[usize::from(rank)] {
        return Err(invalid(format!("duplicate {:?} master-page child", kind)));
    }
    active.seen[usize::from(rank)] = true;
    active.last_rank = Some(rank);
    Ok(())
}

fn push_child(
    xml: &str,
    pages: &mut [Master],
    active: &mut ActiveMaster,
    kind: ChildKind,
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
    children.push(Child {
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
    page: &Master,
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

pub(crate) fn validate_schema(xml: &str, pages: &mut [Master]) -> Result<()> {
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

struct MasterBuilder {
    page: Master,
    start: usize,
    depth: usize,
}

struct RegionBuilder {
    kind: Kind,
    start: usize,
    depth: usize,
    text: String,
    expanded_spaces: usize,
}

pub fn read(xml: &str) -> Result<Vec<Master>> {
    // quick-xml strips a UTF-8 BOM and reports positions relative to the
    // stripped text, so slice against the same view.
    let xml = xml.strip_prefix('\u{FEFF}').unwrap_or(xml);
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut pages = Vec::new();
    let mut master: Option<MasterBuilder> = None;
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
                master = Some(MasterBuilder {
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
                    && let Some(kind) = Kind::parse(element.local_name().as_ref())
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
                    && Kind::parse(element.local_name().as_ref()).is_some() =>
            {
                let kind = Kind::parse(element.local_name().as_ref()).unwrap();
                let master = master.as_mut().expect("checked master page");
                if master.page.region(kind).is_some() {
                    return Err(Error::InvalidFormat(format!(
                        "duplicate {kind:?} in master page '{}'",
                        master.page.name
                    )));
                }
                master.page.regions.push(Region {
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
                        master.page.regions.push(Region {
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
    validate_schema(xml, &mut pages)?;
    let mut structured = parse(xml)?;
    for page in &mut pages {
        for region in &mut page.regions {
            region.blocks = structured
                .remove(&(page.name.clone(), region.kind))
                .unwrap_or_default();
        }
    }
    Ok(pages)
}

fn parse_master_page(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Master> {
    let name = style_attr(reader, element, b"name")?.ok_or_else(|| {
        Error::InvalidFormat("style:master-page is missing style:name".to_string())
    })?;
    Ok(Master {
        name,
        display_name: style_attr(reader, element, b"display-name")?,
        page_layout_name: style_attr(reader, element, b"page-layout-name")?,
        drawing_style_name: namespaced_attr(reader, element, DRAW_NAMESPACE, b"style-name")?,
        next_style_name: style_attr(reader, element, b"next-style-name")?,
        regions: Vec::new(),
        children: Vec::new(),
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
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| Error::XmlError(error.to_string()))?;
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

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";

    #[test]
    fn reads_master_regions_and_lossless_attributes() {
        let xml = format!(
            r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:s="{STYLE}" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:master-styles><s:master-page s:name="Standard" s:display-name="Default &amp; Main" s:page-layout-name="pm1"><s:header><t:p>Page <t:page-number/></t:p><t:p>A<t:s t:c="2"/>B<t:tab/>C<t:line-break/>D</t:p></s:header><s:footer/></s:master-page></o:master-styles></o:document-styles>"#
        );
        let pages = read(&xml).unwrap();
        assert_eq!(pages.len(), 1);
        assert_eq!(pages[0].name, "Standard");
        assert_eq!(pages[0].display_name.as_deref(), Some("Default & Main"));
        assert_eq!(
            pages[0].region(Kind::Header).unwrap().text,
            "Page \nA  B\tC\nD"
        );
        assert!(
            pages[0]
                .region(Kind::Header)
                .unwrap()
                .xml
                .contains("page-number")
        );
        assert!(pages[0].region(Kind::Footer).is_some());
    }

    #[test]
    fn rejects_duplicate_regions_and_unknown_direct_children() {
        let duplicate = format!(
            r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:s="{STYLE}"><o:master-styles><s:master-page s:name="A" s:page-layout-name="pm1"><s:header/><s:header/></s:master-page></o:master-styles></o:document-styles>"#
        );
        assert!(read(&duplicate).is_err());

        let unknown = format!(
            r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:s="{STYLE}"><o:master-styles><s:master-page s:name="A" s:page-layout-name="pm1"><x:foreign xmlns:x="urn:example"/></s:master-page></o:master-styles></o:document-styles>"#
        );
        assert!(read(&unknown).is_err());
    }
}
