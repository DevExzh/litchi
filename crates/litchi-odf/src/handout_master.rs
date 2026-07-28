//! ODF `style:handout-master`: the handout master page of a presentation.
//!
//! Every LibreOffice Impress document carries exactly one handout master in
//! `styles.xml`, describing the page layout and referenced presentation page
//! layout used when printing handouts. The model is inert: shape children are
//! preserved verbatim and never laid out or rendered.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
/// Largest accepted `styles.xml` or flat-document input.
const MAX_XML_BYTES: usize = 64 * 1_048_576;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

/// Offset just past the open tag of an element fragment, skipping quoted
/// attribute values that may contain `>`.
fn open_tag_end(fragment: &str) -> Result<usize> {
    let mut quote = None;
    for (index, ch) in fragment.char_indices() {
        match (quote, ch) {
            (None, '\'' | '"') => quote = Some(ch),
            (Some(active), c) if c == active => quote = None,
            (None, '>') => return Ok(index + 1),
            _ => {},
        }
    }
    Err(invalid("handout-master open tag is unterminated"))
}

/// The handout master page of a presentation document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HandoutMaster {
    /// Required `style:page-layout-name` of the handout page layout.
    pub page_layout_name: String,
    /// Optional `presentation:presentation-page-layout-name`.
    pub presentation_page_layout_name: Option<String>,
    /// Optional `draw:style-name` of the handout drawing style.
    pub drawing_style_name: Option<String>,
    /// Optional `presentation:use-header-name` header declaration reference.
    pub use_header_name: Option<String>,
    /// Optional `presentation:use-footer-name` footer declaration reference.
    pub use_footer_name: Option<String>,
    /// Optional `presentation:use-date-time-name` date-time declaration
    /// reference.
    pub use_date_time_name: Option<String>,
    /// Verbatim XML of the shape children, in document order.
    pub shapes_xml: String,
    /// Verbatim XML of the complete `style:handout-master` element.
    pub xml: String,
}

fn namespaced_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    local_name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("handout-master attribute error: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == expected_namespace)
            && local.as_ref() == local_name
        {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| invalid(format!("handout-master attribute value: {error}")))?;
            return Ok(Some(value.into_owned()));
        }
    }
    Ok(None)
}

/// Parse the `style:handout-master` element from a `styles.xml` document or a
/// flat OpenDocument, returning `None` when none is present.
pub fn parse_handout_master(xml: &str) -> Result<Option<HandoutMaster>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("handout-master input exceeds 64 MiB"));
    }
    // quick-xml strips a UTF-8 BOM and reports positions relative to the
    // stripped text, so slice against the same view.
    let xml = xml.strip_prefix('\u{FEFF}').unwrap_or(xml);
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut active: Option<(HandoutMaster, usize, usize)> = None;
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("handout-master parsing error: {error}")))?;
        let style_element =
            matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == STYLE);
        let event = event.into_owned();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("handout-master nesting overflow"))?;
                if style_element && element.local_name().as_ref() == b"handout-master" {
                    if active.is_some() {
                        return Err(invalid("nested style:handout-master element"));
                    }
                    let page_layout_name = namespaced_attr(&reader, &element, STYLE, b"page-layout-name")?
                        .ok_or_else(|| {
                            invalid("style:handout-master is missing style:page-layout-name")
                        })?;
                    active = Some((
                        HandoutMaster {
                            page_layout_name,
                            presentation_page_layout_name: namespaced_attr(
                                &reader,
                                &element,
                                PRESENTATION,
                                b"presentation-page-layout-name",
                            )?,
                            drawing_style_name: namespaced_attr(
                                &reader,
                                &element,
                                DRAW,
                                b"style-name",
                            )?,
                            use_header_name: namespaced_attr(
                                &reader,
                                &element,
                                PRESENTATION,
                                b"use-header-name",
                            )?,
                            use_footer_name: namespaced_attr(
                                &reader,
                                &element,
                                PRESENTATION,
                                b"use-footer-name",
                            )?,
                            use_date_time_name: namespaced_attr(
                                &reader,
                                &element,
                                PRESENTATION,
                                b"use-date-time-name",
                            )?,
                            shapes_xml: String::new(),
                            xml: String::new(),
                        },
                        event_start,
                        depth,
                    ));
                }
            },
            Event::Empty(element) => {
                if style_element && element.local_name().as_ref() == b"handout-master" {
                    if active.is_some() {
                        return Err(invalid("duplicate style:handout-master element"));
                    }
                    let page_layout_name = namespaced_attr(&reader, &element, STYLE, b"page-layout-name")?
                        .ok_or_else(|| {
                            invalid("style:handout-master is missing style:page-layout-name")
                        })?;
                    return Ok(Some(HandoutMaster {
                        page_layout_name,
                        presentation_page_layout_name: namespaced_attr(
                            &reader,
                            &element,
                            PRESENTATION,
                            b"presentation-page-layout-name",
                        )?,
                        drawing_style_name: namespaced_attr(&reader, &element, DRAW, b"style-name")?,
                        use_header_name: namespaced_attr(
                            &reader,
                            &element,
                            PRESENTATION,
                            b"use-header-name",
                        )?,
                        use_footer_name: namespaced_attr(
                            &reader,
                            &element,
                            PRESENTATION,
                            b"use-footer-name",
                        )?,
                        use_date_time_name: namespaced_attr(
                            &reader,
                            &element,
                            PRESENTATION,
                            b"use-date-time-name",
                        )?,
                        shapes_xml: String::new(),
                        xml: xml[event_start..event_end].to_string(),
                    }));
                }
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("handout-master nesting underflow"))?;
                if let Some((master, start, start_depth)) = active.take() {
                    if depth + 1 == start_depth
                        && style_element
                        && element.local_name().as_ref() == b"handout-master"
                    {
                        let mut master = master;
                        master.shapes_xml = xml[open_tag_end(&xml[start..event_start])?..]
                            .to_string();
                        master.xml = xml[start..event_end].to_string();
                        return Ok(Some(master));
                    }
                    active = Some((master, start, start_depth));
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(None)
}

impl crate::OpenDocumentPackage {
    /// The handout master page declared in packaged `styles.xml`, when present.
    pub fn handout_master(&self) -> Result<Option<HandoutMaster>> {
        match self.styles_xml()? {
            Some(xml) => parse_handout_master(&xml),
            None => Ok(None),
        }
    }
}

impl crate::FlatOpenDocument {
    /// The handout master page declared in a flat OpenDocument, when present.
    pub fn handout_master(&self) -> Result<Option<HandoutMaster>> {
        parse_handout_master(self.xml())
    }
}

impl crate::OpenDocumentPackage {
    /// The unnamed fallback page layout (`style:default-page-layout`) declared
    /// in packaged `styles.xml`, when present.
    pub fn default_page_layout(&self) -> Result<Option<crate::odt::PageLayout>> {
        match self.styles_xml()? {
            Some(xml) => crate::odt::page_layout::parse_default_page_layout(&xml),
            None => Ok(None),
        }
    }
}

impl crate::FlatOpenDocument {
    /// The unnamed fallback page layout (`style:default-page-layout`) declared
    /// in a flat OpenDocument, when present.
    pub fn default_page_layout(&self) -> Result<Option<crate::odt::PageLayout>> {
        crate::odt::page_layout::parse_default_page_layout(self.xml())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const STYLES: &str = concat!(
        r#"<?xml version="1.0"?><office:document-styles "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
        r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
        r#"xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" "#,
        r#"office:version="1.3"><office:master-styles>"#,
        r#"<style:handout-master presentation:presentation-page-layout-name="AL0T26" "#,
        r#"style:page-layout-name="PM0" draw:style-name="Mdp2" "#,
        r#"presentation:use-header-name="hdr1">"#,
        r#"<draw:rect draw:style-name="gr1" draw:name="Shape1"/>"#,
        r#"</style:handout-master></office:master-styles></office:document-styles>"#,
    );

    #[test]
    fn parses_handout_master_attributes_and_shapes() {
        let master = parse_handout_master(STYLES).unwrap().unwrap();
        assert_eq!(master.page_layout_name, "PM0");
        assert_eq!(master.presentation_page_layout_name.as_deref(), Some("AL0T26"));
        assert_eq!(master.drawing_style_name.as_deref(), Some("Mdp2"));
        assert_eq!(master.use_header_name.as_deref(), Some("hdr1"));
        assert_eq!(master.use_footer_name, None);
        assert!(master.shapes_xml.contains("draw:rect"));
        assert!(master.xml.starts_with("<style:handout-master"));
        assert!(master.xml.ends_with("</style:handout-master>"));
    }

    #[test]
    fn returns_none_without_a_handout_master() {
        let xml = concat!(
            r#"<?xml version="1.0"?><office:document-styles "#,
            r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
            r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
            r#"office:version="1.3"><office:master-styles/></office:document-styles>"#,
        );
        assert_eq!(parse_handout_master(xml).unwrap(), None);
    }

    #[test]
    fn rejects_a_missing_page_layout_name() {
        let xml = STYLES.replace("style:page-layout-name=\"PM0\" ", "");
        assert!(parse_handout_master(&xml).is_err());
    }
}
