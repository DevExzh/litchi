//! Page-layout styles from ODT `styles.xml`.

use std::collections::HashSet;

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";

/// The pages to which a page layout applies.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum PageUsage {
    /// Apply the layout to all pages. This is the ODF default.
    #[default]
    All,
    /// Apply the layout to left pages.
    Left,
    /// Apply the layout to right pages.
    Right,
    /// Mirror inner and outer page properties on facing pages.
    Mirrored,
}

impl PageUsage {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "all" => Ok(Self::All),
            "left" => Ok(Self::Left),
            "right" => Ok(Self::Right),
            "mirrored" => Ok(Self::Mirrored),
            _ => Err(Error::InvalidFormat(format!(
                "invalid style:page-usage '{value}'"
            ))),
        }
    }

    /// Return the ODF lexical value.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::All => "all",
            Self::Left => "left",
            Self::Right => "right",
            Self::Mirrored => "mirrored",
        }
    }
}

/// One expanded-name attribute on `style:page-layout-properties`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PageLayoutAttribute {
    namespace_uri: Option<String>,
    local_name: String,
    qualified_name: String,
    value: String,
}

impl PageLayoutAttribute {
    /// Return the resolved namespace URI, or `None` for an unqualified attribute.
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    /// Return the local name without a namespace prefix.
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Return the qualified name exactly as written in `styles.xml`.
    pub fn qualified_name(&self) -> &str {
        &self.qualified_name
    }

    /// Return the decoded attribute value.
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// Page geometry and printing properties with losslessly retained child XML.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PageLayoutProperties {
    attributes: Vec<PageLayoutAttribute>,
    /// Typed multi-column layout, if present.
    pub columns: Option<crate::style::columns::Columns>,
    /// Typed footnote separator, if present.
    pub footnote_separator: Option<crate::footnote_separator::Separator>,
    /// The exact `style:page-layout-properties` element, including background,
    /// columns, and footnote-separator children.
    pub xml: String,
}

impl PageLayoutProperties {
    /// Return every property attribute in source order.
    pub fn attributes(&self) -> &[PageLayoutAttribute] {
        &self.attributes
    }

    /// Find a property by expanded XML name.
    pub fn attribute(&self, namespace_uri: Option<&str>, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace_uri() == namespace_uri && attribute.local_name == local_name
            })
            .map(PageLayoutAttribute::value)
    }
}

/// An automatic page-layout style referenced by an ODT master page.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PageLayout {
    pub name: String,
    pub page_usage: PageUsage,
    pub properties: Option<PageLayoutProperties>,
    /// Exact `style:header-style` XML, including header/footer properties.
    pub header_style_xml: Option<String>,
    /// Typed standard properties from `style:header-style`.
    pub header_properties: Option<crate::header_footer::properties::StyleProperties>,
    /// Exact `style:footer-style` XML, including header/footer properties.
    pub footer_style_xml: Option<String>,
    /// Typed standard properties from `style:footer-style`.
    pub footer_properties: Option<crate::header_footer::properties::StyleProperties>,
    /// Exact `style:page-layout` element bytes.
    pub xml: String,
}

struct PageLayoutBuilder {
    layout: PageLayout,
    start: usize,
    depth: usize,
    child: Option<ChildCapture>,
}

struct ChildCapture {
    kind: ChildKind,
    start: usize,
    depth: usize,
    attributes: Vec<PageLayoutAttribute>,
}

#[derive(Clone, Copy)]
enum ChildKind {
    Properties,
    HeaderStyle,
    FooterStyle,
}

pub(crate) fn parse_page_layouts(xml: &str) -> Result<Vec<PageLayout>> {
    scan_page_layouts(xml, b"page-layout", true)
}

/// Parse the optional `style:default-page-layout` of a document: the unnamed
/// fallback page layout with the same children as `style:page-layout`.
pub(crate) fn parse_default_page_layout(xml: &str) -> Result<Option<PageLayout>> {
    let mut layouts = scan_page_layouts(xml, b"default-page-layout", false)?;
    if layouts.len() > 1 {
        return Err(Error::InvalidFormat(
            "document contains more than one style:default-page-layout".to_string(),
        ));
    }
    Ok(layouts.pop())
}

fn scan_page_layouts(
    xml: &str,
    local_name: &'static [u8],
    require_name: bool,
) -> Result<Vec<PageLayout>> {
    // quick-xml strips a UTF-8 BOM and reports positions relative to the
    // stripped text, so slice against the same view.
    let xml = xml.strip_prefix('\u{FEFF}').unwrap_or(xml);
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut layouts = Vec::new();
    let mut active: Option<PageLayoutBuilder> = None;

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
                    && element.local_name().as_ref() == local_name =>
            {
                active = Some(PageLayoutBuilder {
                    layout: parse_page_layout(&reader, &element, require_name)?,
                    start: event_start,
                    depth: 1,
                    child: None,
                });
            },
            Event::Empty(element)
                if active.is_none()
                    && style_element
                    && element.local_name().as_ref() == local_name =>
            {
                let mut layout = parse_page_layout(&reader, &element, require_name)?;
                layout.xml = xml[event_start..event_end].to_string();
                push_layout(&mut layouts, layout)?;
            },
            Event::Start(element) if active.is_some() => {
                let builder = active.as_mut().expect("checked page layout");
                let direct_child = builder.depth == 1;
                if let Some(child) = builder.child.as_mut() {
                    child.depth += 1;
                } else if direct_child
                    && style_element
                    && let Some(kind) = child_kind(element.local_name().as_ref())
                {
                    ensure_child_absent(&builder.layout, kind)?;
                    let attributes = if matches!(kind, ChildKind::Properties) {
                        parse_property_attributes(&reader, &element)?
                    } else {
                        Vec::new()
                    };
                    builder.child = Some(ChildCapture {
                        kind,
                        start: event_start,
                        depth: 1,
                        attributes,
                    });
                }
                builder.depth += 1;
            },
            Event::Empty(element) if active.is_some() => {
                let builder = active.as_mut().expect("checked page layout");
                if builder.depth == 1
                    && builder.child.is_none()
                    && style_element
                    && let Some(kind) = child_kind(element.local_name().as_ref())
                {
                    ensure_child_absent(&builder.layout, kind)?;
                    let attributes = if matches!(kind, ChildKind::Properties) {
                        parse_property_attributes(&reader, &element)?
                    } else {
                        Vec::new()
                    };
                    store_child(
                        &mut builder.layout,
                        kind,
                        attributes,
                        xml[event_start..event_end].to_string(),
                    )?;
                }
            },
            Event::End(element) if active.is_some() => {
                let builder = active.as_mut().expect("checked page layout");
                if let Some(child) = builder.child.as_mut() {
                    child.depth = child.depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid page-layout child nesting".to_string())
                    })?;
                    if child.depth == 0 {
                        let child = builder.child.take().expect("checked page-layout child");
                        store_child(
                            &mut builder.layout,
                            child.kind,
                            child.attributes,
                            xml[child.start..event_end].to_string(),
                        )?;
                    }
                }
                builder.depth = builder.depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid page-layout nesting".to_string())
                })?;
                if builder.depth == 0 {
                    if !style_element || element.local_name().as_ref() != local_name {
                        return Err(Error::InvalidFormat(
                            "malformed page-layout element".to_string(),
                        ));
                    }
                    let mut finished = active.take().expect("checked page layout");
                    finished.layout.xml = xml[finished.start..event_end].to_string();
                    push_layout(&mut layouts, finished.layout)?;
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if active.is_some() {
        return Err(Error::InvalidFormat(
            "unterminated page-layout element".to_string(),
        ));
    }
    Ok(layouts)
}

pub(crate) fn set_page_layout_xml(
    styles_xml: &str,
    page_layout_name: &str,
    page_layout_xml: &str,
) -> Result<String> {
    validate_page_layout_xml(page_layout_name, page_layout_xml)?;
    if !parse_page_layouts(styles_xml)?
        .iter()
        .any(|layout| layout.name == page_layout_name)
    {
        return Err(Error::InvalidFormat(format!(
            "page layout '{page_layout_name}' does not exist"
        )));
    }
    let (start, end) = find_page_layout(styles_xml, page_layout_name)?.ok_or_else(|| {
        Error::InvalidFormat(format!("page layout '{page_layout_name}' does not exist"))
    })?;
    let mut output =
        String::with_capacity(styles_xml.len() - (end - start) + page_layout_xml.len());
    output.push_str(&styles_xml[..start]);
    output.push_str(page_layout_xml);
    output.push_str(&styles_xml[end..]);
    Ok(output)
}

fn validate_page_layout_xml(expected_name: &str, page_layout_xml: &str) -> Result<()> {
    let wrapper = format!(
        "<office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"{}\"><office:automatic-styles>{page_layout_xml}</office:automatic-styles></office:document-styles>",
        String::from_utf8_lossy(STYLE_NAMESPACE),
    );
    let layouts = parse_page_layouts(&wrapper)?;
    if layouts.len() != 1 || layouts[0].name != expected_name || layouts[0].xml != page_layout_xml {
        return Err(Error::InvalidFormat(format!(
            "page-layout XML must be exactly one style:page-layout named '{expected_name}'"
        )));
    }
    Ok(())
}

fn find_page_layout(xml: &str, expected_name: &str) -> Result<Option<(usize, usize)>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut active: Option<(usize, usize)> = None;
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
                    && element.local_name().as_ref() == b"page-layout"
                    && style_attr(&reader, &element, b"name")?.as_deref()
                        == Some(expected_name) =>
            {
                active = Some((event_start, 1));
            },
            Event::Empty(element)
                if active.is_none()
                    && style_element
                    && element.local_name().as_ref() == b"page-layout"
                    && style_attr(&reader, &element, b"name")?.as_deref()
                        == Some(expected_name) =>
            {
                return Ok(Some((event_start, event_end)));
            },
            Event::Start(_) if active.is_some() => active.as_mut().unwrap().1 += 1,
            Event::End(_) if active.is_some() => {
                let current = active.as_mut().unwrap();
                current.1 = current.1.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid page-layout nesting".to_string())
                })?;
                if current.1 == 0 {
                    return Ok(Some((current.0, event_end)));
                }
            },
            Event::Eof => return Ok(None),
            _ => {},
        }
        buffer.clear();
    }
}

fn parse_page_layout(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    require_name: bool,
) -> Result<PageLayout> {
    let name = match style_attr(reader, element, b"name")? {
        Some(name) => name,
        None if require_name => {
            return Err(Error::InvalidFormat(
                "style:page-layout is missing style:name".to_string(),
            ));
        },
        None => String::new(),
    };
    let page_usage = style_attr(reader, element, b"page-usage")?
        .as_deref()
        .map(PageUsage::parse)
        .transpose()?
        .unwrap_or_default();
    Ok(PageLayout {
        name,
        page_usage,
        properties: None,
        header_style_xml: None,
        header_properties: None,
        footer_style_xml: None,
        footer_properties: None,
        xml: String::new(),
    })
}

fn parse_property_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Vec<PageLayoutAttribute>> {
    if element.attributes().count() > 256 {
        return Err(Error::InvalidFormat(
            "page-layout-properties exceeds 256 attributes".to_string(),
        ));
    }
    let mut attributes = Vec::with_capacity(element.attributes().count());
    let mut expanded_names = HashSet::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid page-layout property: {error}"))
        })?;
        let qualified_name = attribute.key.as_ref();
        if qualified_name == b"xmlns" || qualified_name.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = match namespace {
            ResolveResult::Unbound => None,
            ResolveResult::Bound(Namespace(uri)) => Some(
                std::str::from_utf8(uri)
                    .map_err(|_| {
                        Error::InvalidFormat("non-UTF-8 property namespace URI".to_string())
                    })?
                    .to_string(),
            ),
            ResolveResult::Unknown(prefix) => {
                return Err(Error::InvalidFormat(format!(
                    "unknown page-layout property prefix '{}'",
                    String::from_utf8_lossy(&prefix)
                )));
            },
        };
        let local_name = std::str::from_utf8(local.as_ref())
            .map_err(|_| Error::InvalidFormat("non-UTF-8 property name".to_string()))?
            .to_string();
        if !expanded_names.insert((namespace_uri.clone(), local_name.clone())) {
            return Err(Error::InvalidFormat(format!(
                "duplicate page-layout property '{local_name}'"
            )));
        }
        let qualified_name = std::str::from_utf8(qualified_name)
            .map_err(|_| Error::InvalidFormat("non-UTF-8 qualified property name".to_string()))?
            .to_string();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid page-layout property value: {error}"))
            })?
            .into_owned();
        attributes.push(PageLayoutAttribute {
            namespace_uri,
            local_name,
            qualified_name,
            value,
        });
    }
    Ok(attributes)
}

fn child_kind(local_name: &[u8]) -> Option<ChildKind> {
    match local_name {
        b"page-layout-properties" => Some(ChildKind::Properties),
        b"header-style" => Some(ChildKind::HeaderStyle),
        b"footer-style" => Some(ChildKind::FooterStyle),
        _ => None,
    }
}

fn ensure_child_absent(layout: &PageLayout, kind: ChildKind) -> Result<()> {
    let duplicate = match kind {
        ChildKind::Properties => layout.properties.is_some(),
        ChildKind::HeaderStyle => layout.header_style_xml.is_some(),
        ChildKind::FooterStyle => layout.footer_style_xml.is_some(),
    };
    if duplicate {
        return Err(Error::InvalidFormat(format!(
            "duplicate style:{} in page layout '{}'",
            match kind {
                ChildKind::Properties => "page-layout-properties",
                ChildKind::HeaderStyle => "header-style",
                ChildKind::FooterStyle => "footer-style",
            },
            layout.name
        )));
    }
    Ok(())
}

fn store_child(
    layout: &mut PageLayout,
    kind: ChildKind,
    attributes: Vec<PageLayoutAttribute>,
    xml: String,
) -> Result<()> {
    match kind {
        ChildKind::Properties => {
            let mut parsed = crate::style::columns::parse_page_layout_property_columns(&xml)?;
            if parsed.len() > 1 {
                return Err(Error::InvalidFormat(
                    "page-layout-properties has multiple style:columns children".to_string(),
                ));
            }
            let mut separators =
                crate::footnote_separator::parse_page_layout_property_footnote_separators(&xml)?;
            if separators.len() > 1 {
                return Err(Error::InvalidFormat(
                    "page-layout-properties has multiple style:footnote-sep children".to_string(),
                ));
            }
            layout.properties = Some(PageLayoutProperties {
                attributes,
                columns: parsed.pop(),
                footnote_separator: separators.pop(),
                xml,
            });
        },
        ChildKind::HeaderStyle => {
            layout.header_properties =
                crate::header_footer::properties::parse_region_properties(&xml)?;
            layout.header_style_xml = Some(xml);
        },
        ChildKind::FooterStyle => {
            layout.footer_properties =
                crate::header_footer::properties::parse_region_properties(&xml)?;
            layout.footer_style_xml = Some(xml);
        },
    }
    Ok(())
}

fn push_layout(layouts: &mut Vec<PageLayout>, layout: PageLayout) -> Result<()> {
    if layouts.iter().any(|existing| existing.name == layout.name) {
        return Err(Error::InvalidFormat(format!(
            "duplicate page layout '{}'",
            layout.name
        )));
    }
    layouts.push(layout);
    Ok(())
}

fn style_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid page-layout attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if bound_to(&namespace, STYLE_NAMESPACE) && local.as_ref() == local_name {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid page-layout attribute: {error}"))
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

    const FO_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";

    #[test]
    fn parses_page_layouts_losslessly_with_all_property_namespaces() {
        let xml = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:x="urn:example:extension"><o:automatic-styles><s:page-layout s:name="pm1" s:page-usage="mirrored"><s:page-layout-properties f:page-width="21cm" f:page-height="29.7cm" f:margin="2cm" s:print-orientation="portrait" x:bleed="3mm"><s:background-image/></s:page-layout-properties><s:header-style><s:header-footer-properties f:min-height="1cm"/></s:header-style><s:footer-style><s:header-footer-properties f:min-height="1.2cm"/></s:footer-style></s:page-layout><s:page-layout s:name="empty"/></o:automatic-styles></o:document-styles>"#;
        let layouts = parse_page_layouts(xml).unwrap();

        assert_eq!(layouts.len(), 2);
        let layout = &layouts[0];
        assert_eq!(layout.name, "pm1");
        assert_eq!(layout.page_usage, PageUsage::Mirrored);
        assert!(layout.xml.starts_with("<s:page-layout "));
        assert!(layout.xml.ends_with("</s:page-layout>"));
        let properties = layout.properties.as_ref().unwrap();
        assert_eq!(
            properties.attribute(Some(FO_NAMESPACE), "page-width"),
            Some("21cm")
        );
        assert_eq!(
            properties.attribute(Some(STYLE_NAMESPACE_STR), "print-orientation"),
            Some("portrait")
        );
        assert_eq!(
            properties.attribute(Some("urn:example:extension"), "bleed"),
            Some("3mm")
        );
        assert_eq!(properties.attributes()[0].qualified_name(), "f:page-width");
        assert!(properties.xml.contains("<s:background-image/>"));
        assert!(
            layout
                .header_style_xml
                .as_deref()
                .unwrap()
                .contains("f:min-height=\"1cm\"")
        );
        assert!(layout.footer_style_xml.is_some());
        assert_eq!(layouts[1].page_usage, PageUsage::All);
        assert_eq!(layouts[1].xml, "<s:page-layout s:name=\"empty\"/>");
    }

    #[test]
    fn rejects_invalid_or_ambiguous_page_layouts() {
        let prefix = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><o:automatic-styles>"#;
        let suffix = "</o:automatic-styles></o:document-styles>";
        assert!(parse_page_layouts(&format!("{prefix}<s:page-layout/>{suffix}")).is_err());
        assert!(
            parse_page_layouts(&format!(
                "{prefix}<s:page-layout s:name=\"A\" s:page-usage=\"both\"/>{suffix}"
            ))
            .is_err()
        );
        assert!(
            parse_page_layouts(&format!(
                "{prefix}<s:page-layout s:name=\"A\"/><s:page-layout s:name=\"A\"/>{suffix}"
            ))
            .is_err()
        );
        assert!(
            parse_page_layouts(&format!(
                "{prefix}<s:page-layout s:name=\"A\"><s:header-style/><s:header-style/></s:page-layout>{suffix}"
            ))
            .is_err()
        );
    }

    #[test]
    fn replaces_one_complete_page_layout_without_rewriting_siblings() {
        let styles = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><o:styles><s:style s:name="keep"/></o:styles><o:automatic-styles><s:page-layout s:name="pm1"/><s:page-layout s:name="pm2" s:page-usage="right"/></o:automatic-styles></o:document-styles>"#;
        let replacement = r#"<x:page-layout xmlns:x="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" x:name="pm1" x:page-usage="mirrored"><x:page-layout-properties f:page-width="21cm" f:margin="2cm"/><x:header-style><x:header-footer-properties f:min-height="1cm"/></x:header-style></x:page-layout>"#;
        let updated = set_page_layout_xml(styles, "pm1", replacement).unwrap();
        let layouts = parse_page_layouts(&updated).unwrap();

        assert_eq!(layouts.len(), 2);
        assert_eq!(layouts[0].xml, replacement);
        assert_eq!(layouts[0].page_usage, PageUsage::Mirrored);
        assert_eq!(
            layouts[1].xml,
            "<s:page-layout s:name=\"pm2\" s:page-usage=\"right\"/>"
        );
        assert!(updated.contains("<s:style s:name=\"keep\"/>"));
        assert!(set_page_layout_xml(styles, "pm1", "not XML").is_err());
        assert!(set_page_layout_xml(styles, "pm1", "<s:page-layout s:name=\"renamed\"/>").is_err());
        assert!(set_page_layout_xml(styles, "missing", replacement).is_err());
    }

    const STYLE_NAMESPACE_STR: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
}
