//! Read-only semantic access to OpenDocument drawings.

use crate::odp::OdpParser;
use crate::{MediaReference, OdfMetadata, OpenDocumentFamily, OpenDocumentPackage, Shape, Slide};
use litchi_core::{Error, Metadata, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::io::Read;
use std::path::Path;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";

/// Exact standard attributes attached to an OpenDocument drawing page.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DrawingPageProperties {
    name: Option<String>,
    draw_id: Option<String>,
    xml_id: Option<String>,
    style_name: Option<String>,
    master_page_name: Option<String>,
    navigation_order: Option<String>,
    presentation_layout_name: Option<String>,
    header_name: Option<String>,
    footer_name: Option<String>,
    date_time_name: Option<String>,
}

impl DrawingPageProperties {
    /// Return `draw:name`.
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Return the legacy `draw:id` identifier.
    pub fn draw_id(&self) -> Option<&str> {
        self.draw_id.as_deref()
    }

    /// Return the namespace-defined `xml:id` identifier.
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }

    /// Return the drawing-page `draw:style-name` reference.
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Return the required-by-schema `draw:master-page-name` reference.
    ///
    /// The value remains optional in the API so legacy or producer-specific
    /// documents that omit it can still be inspected.
    pub fn master_page_name(&self) -> Option<&str> {
        self.master_page_name.as_deref()
    }

    /// Return the exact whitespace-separated `draw:nav-order` IDREFS value.
    pub fn navigation_order(&self) -> Option<&str> {
        self.navigation_order.as_deref()
    }

    /// Return `presentation:presentation-page-layout-name`.
    pub fn presentation_layout_name(&self) -> Option<&str> {
        self.presentation_layout_name.as_deref()
    }

    /// Return `presentation:use-header-name`.
    pub fn header_name(&self) -> Option<&str> {
        self.header_name.as_deref()
    }

    /// Return `presentation:use-footer-name`.
    pub fn footer_name(&self) -> Option<&str> {
        self.footer_name.as_deref()
    }

    /// Return `presentation:use-date-time-name`.
    pub fn date_time_name(&self) -> Option<&str> {
        self.date_time_name.as_deref()
    }
}

/// A page in an OpenDocument drawing.
#[derive(Debug, Clone)]
pub struct DrawingPage {
    properties: DrawingPageProperties,
    page: Slide,
}

impl DrawingPage {
    /// Return the zero-based page index.
    pub fn index(&self) -> usize {
        self.page.index()
    }

    /// Return the exact optional `draw:name` value.
    pub fn name(&self) -> Option<&str> {
        self.properties.name()
    }

    /// Return all exact standard page attributes.
    pub fn properties(&self) -> &DrawingPageProperties {
        &self.properties
    }

    /// Return every top-level drawing shape on the page.
    ///
    /// Unlike the presentation API, text and title frames remain shapes.
    pub fn shapes(&self) -> &[Shape] {
        &self.page.shapes
    }

    /// Compose visible text from this page's shapes, including nested groups.
    pub fn text(&self) -> String {
        self.page.all_text()
    }
}

/// A validated OpenDocument drawing (`.odg`) or drawing template (`.otg`).
///
/// The semantic page model is parsed eagerly. Unmodified saves return the
/// original package bytes exactly, including parts unknown to this library.
pub struct DrawingDocument {
    package: OpenDocumentPackage,
    pages: Vec<DrawingPage>,
}

impl DrawingDocument {
    /// Open and validate an OpenDocument drawing from a path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Read and validate an OpenDocument drawing stream.
    pub fn from_reader(mut reader: impl Read) -> Result<Self> {
        let mut bytes = Vec::new();
        reader.read_to_end(&mut bytes)?;
        Self::from_bytes(bytes)
    }

    /// Validate and parse an OpenDocument drawing from owned package bytes.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = OpenDocumentPackage::from_bytes(bytes)?;
        if package.family() != OpenDocumentFamily::Drawing {
            return Err(Error::InvalidFormat(format!(
                "not an OpenDocument drawing: MIME type is '{}'",
                package.mimetype()
            )));
        }

        let content = package.content_xml()?;
        let properties = validate_drawing_content(&content)?;
        let styles = package.styles_xml()?;
        let parsed = OdpParser::parse_drawing_pages(&content, styles.as_deref())?;
        if parsed.len() != properties.len() {
            return Err(Error::InvalidFormat(
                "drawing page structure changed during semantic parsing".to_string(),
            ));
        }
        let pages = properties
            .into_iter()
            .zip(parsed)
            .map(|(properties, page)| DrawingPage { properties, page })
            .collect();

        Ok(Self { package, pages })
    }

    /// Whether this document is an `.otg` drawing template.
    pub fn is_template(&self) -> bool {
        self.package.is_template()
    }

    /// Return the exact package MIME type.
    pub fn mimetype(&self) -> &str {
        self.package.mimetype()
    }

    /// Return the number of drawing pages.
    pub fn page_count(&self) -> usize {
        self.pages.len()
    }

    /// Return all drawing pages.
    pub fn pages(&self) -> &[DrawingPage] {
        &self.pages
    }

    /// Return a drawing page by zero-based index.
    pub fn page(&self, index: usize) -> Option<&DrawingPage> {
        self.pages.get(index)
    }

    /// Compose visible text from all non-empty drawing pages.
    pub fn text(&self) -> String {
        self.pages
            .iter()
            .map(DrawingPage::text)
            .filter(|text| !text.is_empty())
            .collect::<Vec<_>>()
            .join("\n\n")
    }

    /// Extract common document metadata.
    pub fn metadata(&self) -> Result<Metadata> {
        self.package.metadata()
    }

    /// Extract the complete format-specific OpenDocument metadata model.
    pub fn odf_metadata(&self) -> Result<Option<OdfMetadata>> {
        self.package.odf_metadata()
    }

    /// Read package-contained media without following external references.
    pub fn media_data(&self, media: &MediaReference) -> Result<Option<Vec<u8>>> {
        let Some(path) = media.package_path() else {
            return Ok(None);
        };
        if !self.package.has_file(path)? {
            return Ok(None);
        }
        self.package.get_file(path).map(Some)
    }

    /// Return the exact original package bytes.
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// Clone the exact original package bytes.
    pub fn to_bytes(&self) -> Vec<u8> {
        self.package.to_bytes()
    }

    /// Consume this document and return the exact original package bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }

    /// Save the drawing without reconstructing its package.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.package.save(path)
    }
}

fn validate_drawing_content(xml: &str) -> Result<Vec<DrawingPageProperties>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut body_seen = false;
    let mut body_open = false;
    let mut drawing_seen = false;
    let mut drawing_open = false;
    let mut pages = Vec::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid drawing XML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                let is_office = bound_to(&namespace, OFFICE_NAMESPACE);
                let is_drawing = bound_to(&namespace, DRAW_NAMESPACE);
                validate_start_position(
                    &reader,
                    is_office,
                    is_drawing,
                    element,
                    depth,
                    &mut root_seen,
                    root_closed,
                    &mut body_seen,
                    &mut body_open,
                    &mut drawing_seen,
                    &mut drawing_open,
                    &mut pages,
                )?;
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("drawing XML nesting overflow".to_string())
                })?;
            },
            Event::Empty(ref element) => {
                if depth == 0 {
                    return Err(Error::InvalidFormat(
                        "drawing content root cannot be empty".to_string(),
                    ));
                }
                let mut empty_body_open = body_open;
                let mut empty_drawing_open = drawing_open;
                let is_office = bound_to(&namespace, OFFICE_NAMESPACE);
                let is_drawing = bound_to(&namespace, DRAW_NAMESPACE);
                validate_start_position(
                    &reader,
                    is_office,
                    is_drawing,
                    element,
                    depth,
                    &mut root_seen,
                    root_closed,
                    &mut body_seen,
                    &mut empty_body_open,
                    &mut drawing_seen,
                    &mut empty_drawing_open,
                    &mut pages,
                )?;
            },
            Event::End(ref element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("unexpected drawing XML closing tag".to_string())
                })?;
                if bound_to(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"drawing"
                    && depth == 2
                {
                    drawing_open = false;
                } else if bound_to(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"body"
                    && depth == 1
                {
                    body_open = false;
                }
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(ref text) if depth == 0 && !text.iter().all(u8::is_ascii_whitespace) => {
                return Err(Error::InvalidFormat(
                    "text is not allowed outside the drawing content root".to_string(),
                ));
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(Error::InvalidFormat(
                    "content is not allowed outside the drawing content root".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if !root_seen
        || !root_closed
        || depth != 0
        || !body_seen
        || body_open
        || !drawing_seen
        || drawing_open
    {
        return Err(Error::InvalidFormat(
            "drawing content is missing a complete office:body/office:drawing structure"
                .to_string(),
        ));
    }
    Ok(pages)
}

#[allow(clippy::too_many_arguments)]
fn validate_start_position(
    reader: &NsReader<&[u8]>,
    is_office: bool,
    is_drawing: bool,
    element: &BytesStart<'_>,
    depth: usize,
    root_seen: &mut bool,
    root_closed: bool,
    body_seen: &mut bool,
    body_open: &mut bool,
    drawing_seen: &mut bool,
    drawing_open: &mut bool,
    pages: &mut Vec<DrawingPageProperties>,
) -> Result<()> {
    let local = element.local_name();
    if depth == 0 {
        if *root_seen || root_closed || !is_office || local.as_ref() != b"document-content" {
            return Err(Error::InvalidFormat(
                "drawing content must have one office:document-content root".to_string(),
            ));
        }
        *root_seen = true;
        return Ok(());
    }

    if is_office && local.as_ref() == b"body" {
        if depth != 1 || *body_seen {
            return Err(Error::InvalidFormat(
                "office:body is misplaced or duplicated in drawing content".to_string(),
            ));
        }
        *body_seen = true;
        *body_open = true;
    } else if is_office && local.as_ref() == b"drawing" {
        if depth != 2 || !*body_open || *drawing_seen {
            return Err(Error::InvalidFormat(
                "office:drawing is misplaced or duplicated in drawing content".to_string(),
            ));
        }
        *drawing_seen = true;
        *drawing_open = true;
    } else if is_drawing && local.as_ref() == b"page" {
        if depth != 3 || !*drawing_open {
            return Err(Error::InvalidFormat(
                "draw:page must be a direct child of office:drawing".to_string(),
            ));
        }
        if pages.len() >= 1_000_000 {
            return Err(Error::InvalidFormat(
                "drawing exceeds one million pages".to_string(),
            ));
        }
        pages.push(drawing_page_properties(reader, element)?);
    }
    Ok(())
}

fn drawing_page_properties(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<DrawingPageProperties> {
    if element.attributes().count() > 256 {
        return Err(Error::InvalidFormat(
            "drawing page exceeds 256 attributes".to_string(),
        ));
    }
    let mut properties = DrawingPageProperties::default();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid page attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let target = if bound_to(&namespace, DRAW_NAMESPACE) {
            match local.as_ref() {
                b"name" => Some((&mut properties.name, "draw:name")),
                b"id" => Some((&mut properties.draw_id, "draw:id")),
                b"style-name" => Some((&mut properties.style_name, "draw:style-name")),
                b"master-page-name" => {
                    Some((&mut properties.master_page_name, "draw:master-page-name"))
                },
                b"nav-order" => Some((&mut properties.navigation_order, "draw:nav-order")),
                _ => None,
            }
        } else if bound_to(&namespace, PRESENTATION_NAMESPACE) {
            match local.as_ref() {
                b"presentation-page-layout-name" => Some((
                    &mut properties.presentation_layout_name,
                    "presentation:presentation-page-layout-name",
                )),
                b"use-header-name" => {
                    Some((&mut properties.header_name, "presentation:use-header-name"))
                },
                b"use-footer-name" => {
                    Some((&mut properties.footer_name, "presentation:use-footer-name"))
                },
                b"use-date-time-name" => Some((
                    &mut properties.date_time_name,
                    "presentation:use-date-time-name",
                )),
                _ => None,
            }
        } else if bound_to(&namespace, XML_NAMESPACE) && local.as_ref() == b"id" {
            Some((&mut properties.xml_id, "xml:id"))
        } else {
            None
        };
        if let Some((slot, display_name)) = target {
            if slot.is_some() {
                return Err(Error::InvalidFormat(format!(
                    "drawing page has duplicate {display_name} attributes"
                )));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!(
                        "invalid drawing page attribute {display_name}: {error}"
                    ))
                })?
                .into_owned();
            if value.len() > 1_048_576 {
                return Err(Error::InvalidFormat(
                    "drawing page attribute exceeds 1 MiB".to_string(),
                ));
            }
            *slot = Some(value);
        }
    }
    Ok(properties)
}

fn bound_to(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::constants;
    use crate::core::PackageWriter;
    use crate::{DrawingShapeKind, EnhancedGeometryChildKind};
    use std::io::Cursor;

    fn package(mimetype: &str, content: &str) -> Vec<u8> {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(mimetype).unwrap();
        writer
            .add_file(constants::ODF_CONTENT, content.as_bytes())
            .unwrap();
        writer
            .add_file(
                constants::ODF_META,
                br#"<o:document-meta xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="http://purl.org/dc/elements/1.1/"><o:meta><d:title>Drawing title</d:title></o:meta></o:document-meta>"#,
            )
            .unwrap();
        writer
            .add_file_with_media_type("Pictures/pixel.png", b"PNG", "image/png")
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn drawing_xml() -> &'static str {
        r#"<?xml version="1.0"?>
<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
 xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
 xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0">
 <o:body><o:drawing>
  <d:page d:name="First &amp; Best" d:id="legacy1" xml:id="page1"
   d:style-name="dp1" d:master-page-name="Default" d:nav-order="shape1 shape2"
   p:presentation-page-layout-name="layout1" p:use-header-name="header1"
   p:use-footer-name="footer1" p:use-date-time-name="date1">
   <d:frame d:name="Label" s:x="1cm" s:y="2cm"><d:text-box><t:p>Hello</t:p></d:text-box></d:frame>
   <d:g d:name="Group"><d:path d:name="Curve" s:d="M 0 0 L 1 1"/></d:g>
   <d:custom-shape><d:enhanced-geometry d:type="diamond"><d:equation d:name="f0" d:formula="$0/2"/><d:handle d:handle-position="$0 $1"/></d:enhanced-geometry></d:custom-shape>
  </d:page>
  <d:page/>
 </o:drawing></o:body>
</o:document-content>"#
    }

    #[test]
    fn parses_drawing_pages_shapes_text_geometry_and_metadata() {
        let bytes = package(constants::ODF_DRAWING, drawing_xml());
        let document = DrawingDocument::from_bytes(bytes.clone()).unwrap();
        assert!(!document.is_template());
        assert_eq!(document.page_count(), 2);
        assert_eq!(document.pages()[0].index(), 0);
        assert_eq!(document.pages()[0].name(), Some("First & Best"));
        assert_eq!(document.pages()[1].name(), None);
        let properties = document.pages()[0].properties();
        assert_eq!(properties.draw_id(), Some("legacy1"));
        assert_eq!(properties.xml_id(), Some("page1"));
        assert_eq!(properties.style_name(), Some("dp1"));
        assert_eq!(properties.master_page_name(), Some("Default"));
        assert_eq!(properties.navigation_order(), Some("shape1 shape2"));
        assert_eq!(properties.presentation_layout_name(), Some("layout1"));
        assert_eq!(properties.header_name(), Some("header1"));
        assert_eq!(properties.footer_name(), Some("footer1"));
        assert_eq!(properties.date_time_name(), Some("date1"));
        assert_eq!(document.pages()[0].text(), "Hello");
        assert_eq!(document.text(), "Hello");

        let shapes = document.pages()[0].shapes();
        assert_eq!(shapes.len(), 3);
        assert_eq!(shapes[0].drawing_kind(), Some(DrawingShapeKind::Frame));
        assert_eq!(shapes[0].text, "Hello");
        assert_eq!(shapes[1].drawing_kind(), Some(DrawingShapeKind::Group));
        assert_eq!(shapes[1].children().len(), 1);
        assert_eq!(
            shapes[1].children()[0].drawing_kind(),
            Some(DrawingShapeKind::Path)
        );
        let geometry = shapes[2].enhanced_geometry().unwrap();
        assert_eq!(
            geometry.children()[0].kind(),
            EnhancedGeometryChildKind::Equation
        );
        assert_eq!(
            document.metadata().unwrap().title.as_deref(),
            Some("Drawing title")
        );
        assert_eq!(document.to_bytes(), bytes);
        assert_eq!(document.as_bytes(), bytes);
    }

    #[test]
    fn accepts_drawing_templates_and_readers_losslessly() {
        let bytes = package(constants::ODF_DRAWING_TEMPLATE, drawing_xml());
        let document = DrawingDocument::from_reader(Cursor::new(bytes.clone())).unwrap();
        assert!(document.is_template());
        assert_eq!(document.mimetype(), constants::ODF_DRAWING_TEMPLATE);
        assert_eq!(document.into_bytes(), bytes);
    }

    #[test]
    fn reads_only_package_local_media() {
        let bytes = package(constants::ODF_DRAWING, drawing_xml());
        let document = DrawingDocument::from_bytes(bytes).unwrap();
        let local = MediaReference::new("Pictures/pixel.png").unwrap();
        let external = MediaReference::new("https://example.com/video.mp4").unwrap();
        assert_eq!(document.media_data(&local).unwrap(), Some(b"PNG".to_vec()));
        assert_eq!(document.media_data(&external).unwrap(), None);
    }

    #[test]
    fn rejects_other_families() {
        let bytes = package(constants::ODF_PRESENTATION, drawing_xml());
        assert!(DrawingDocument::from_bytes(bytes).is_err());
    }

    #[test]
    fn rejects_wrong_or_incomplete_drawing_body() {
        for xml in [
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:presentation/></o:body></o:document-content>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:drawing>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:body><o:drawing><o:section><d:page/></o:section></o:drawing></o:body></o:document-content>"#,
        ] {
            let bytes = package(constants::ODF_DRAWING, xml);
            assert!(
                DrawingDocument::from_bytes(bytes).is_err(),
                "accepted {xml}"
            );
        }
    }

    #[test]
    fn rejects_duplicate_family_bodies_and_page_names() {
        let duplicate_body = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:drawing/></o:body><o:body><o:drawing/></o:body></o:document-content>"#;
        assert!(
            DrawingDocument::from_bytes(package(constants::ODF_DRAWING, duplicate_body)).is_err()
        );

        let duplicate_name = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:x="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:body><o:drawing><d:page d:name="one" x:name="two"/></o:drawing></o:body></o:document-content>"#;
        assert!(
            DrawingDocument::from_bytes(package(constants::ODF_DRAWING, duplicate_name)).is_err()
        );
    }
}
