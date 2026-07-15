//! Master-page headers and footers from ODT `styles.xml`.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";

/// One of the six header/footer regions supported by an ODF master page.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum HeaderFooterKind {
    Header,
    HeaderFirst,
    HeaderLeft,
    Footer,
    FooterFirst,
    FooterLeft,
}

impl HeaderFooterKind {
    fn parse(local_name: &[u8]) -> Option<Self> {
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
}

/// Losslessly retained content of one master-page header or footer.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HeaderFooter {
    pub kind: HeaderFooterKind,
    /// The exact element bytes from `styles.xml`, including nested fields and formatting.
    pub xml: String,
    /// Best-effort visible literal text. Dynamic field values remain represented in `xml`.
    pub text: String,
}

/// An ODF master page and all of its header/footer regions.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MasterPage {
    pub name: String,
    pub display_name: Option<String>,
    pub page_layout_name: Option<String>,
    pub next_style_name: Option<String>,
    pub regions: Vec<HeaderFooter>,
}

impl MasterPage {
    /// Return a particular header/footer region when it exists.
    pub fn region(&self, kind: HeaderFooterKind) -> Option<&HeaderFooter> {
        self.regions.iter().find(|region| region.kind == kind)
    }
}

struct MasterPageBuilder {
    page: MasterPage,
    depth: usize,
}

struct RegionBuilder {
    kind: HeaderFooterKind,
    start: usize,
    depth: usize,
    text: String,
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
                    depth: 1,
                });
            },
            Event::Empty(element)
                if style_element && element.local_name().as_ref() == b"master-page" =>
            {
                pages.push(parse_master_page(&reader, &element)?);
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
                });
            },
            Event::Empty(element) if region.is_some() && text_element => {
                append_empty_text_element(&reader, &element, &mut region.as_mut().unwrap().text)?;
            },
            Event::Text(value) if region.is_some() => {
                let decoded = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid header text: {error}"))
                    })?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(|error| {
                    Error::InvalidFormat(format!("invalid header character reference: {error}"))
                })?;
                region.as_mut().unwrap().text.push_str(&decoded);
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
                    pages.push(master.take().expect("checked master page").page);
                }
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
    Ok(pages)
}

fn parse_master_page(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<MasterPage> {
    let name = style_attr(reader, element, b"name")?.ok_or_else(|| {
        Error::InvalidFormat("style:master-page is missing style:name".to_string())
    })?;
    Ok(MasterPage {
        name,
        display_name: style_attr(reader, element, b"display-name")?,
        page_layout_name: style_attr(reader, element, b"page-layout-name")?,
        next_style_name: style_attr(reader, element, b"next-style-name")?,
        regions: Vec::new(),
    })
}

fn style_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid master-page attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == STYLE_NAMESPACE)
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
) -> Result<()> {
    match element.local_name().as_ref() {
        b"s" => {
            let count = style_independent_text_count(reader, element)?.unwrap_or(1);
            if count > 1_000_000 {
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_all_master_page_regions_losslessly_with_arbitrary_prefixes() {
        let xml = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:master-styles><s:master-page s:name="Standard" s:display-name="Default &amp; Main" s:page-layout-name="pm1" s:next-style-name="Next"><s:header><t:p>Page <t:page-number/></t:p><t:p>A<t:s t:c="2"/>B<t:tab/>C<t:line-break/>D</t:p></s:header><s:header-first><t:p>First</t:p></s:header-first><s:header-left><t:p>Left</t:p></s:header-left><s:footer><t:p>Footer</t:p></s:footer><s:footer-first><t:p>First footer</t:p></s:footer-first><s:footer-left><t:p>Left footer</t:p></s:footer-left></s:master-page><s:master-page s:name="Empty"/></o:master-styles></o:document-styles>"#;
        let pages = parse_master_pages(xml).unwrap();
        assert_eq!(pages.len(), 2);
        assert_eq!(pages[0].name, "Standard");
        assert_eq!(pages[0].display_name.as_deref(), Some("Default & Main"));
        assert_eq!(pages[0].page_layout_name.as_deref(), Some("pm1"));
        assert_eq!(pages[0].next_style_name.as_deref(), Some("Next"));
        assert_eq!(pages[0].regions.len(), 6);
        let header = pages[0].region(HeaderFooterKind::Header).unwrap();
        assert_eq!(header.text, "Page \nA  B\tC\nD");
        assert!(header.xml.starts_with("<s:header>"));
        assert!(header.xml.contains("<t:page-number/>"));
        assert_eq!(pages[1].name, "Empty");
        assert!(pages[1].regions.is_empty());
    }

    #[test]
    fn rejects_duplicate_regions_and_missing_master_names() {
        let duplicate = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><o:master-styles><s:master-page s:name="A"><s:header/><s:header/></s:master-page></o:master-styles></o:document-styles>"#;
        assert!(parse_master_pages(duplicate).is_err());
        let missing = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><o:master-styles><s:master-page/></o:master-styles></o:document-styles>"#;
        assert!(parse_master_pages(missing).is_err());
    }
}
