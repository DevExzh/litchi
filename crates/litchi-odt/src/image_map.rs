//! ODF `draw:image-map`: client-side image maps with clickable areas.
//!
//! An image map attaches link targets to geometric areas of an image frame.
//! Everything here is inert: link targets are stored verbatim and are never
//! resolved, followed, fetched, or rendered, and event-listener content is
//! preserved without interpretation.

use litchi_core::{Error, Result};
use quick_xml::{
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const SVG: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
/// Largest accepted input document.
const MAX_XML_BYTES: usize = 64 * 1_048_576;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn namespaced_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    local_name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("image-map attribute error: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == expected_namespace)
            && local.as_ref() == local_name
        {
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| invalid(format!("image-map attribute value: {error}")))?;
            return Ok(Some(value.into_owned()));
        }
    }
    Ok(None)
}

fn required_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local_name: &[u8],
) -> Result<String> {
    namespaced_attr(reader, element, namespace, local_name)?.ok_or_else(|| {
        invalid(format!(
            "draw:{} is missing its {} attribute",
            String::from_utf8_lossy(element.local_name().as_ref()),
            String::from_utf8_lossy(local_name)
        ))
    })
}

fn bool_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local_name: &[u8],
) -> Result<bool> {
    match namespaced_attr(reader, element, namespace, local_name)?.as_deref() {
        None | Some("false") => Ok(false),
        Some("true") => Ok(true),
        Some(value) => Err(invalid(format!(
            "image-map Boolean attribute {} has value '{value}'",
            String::from_utf8_lossy(local_name)
        ))),
    }
}

/// The geometry of one clickable image-map area.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ImageMapAreaShape {
    /// `draw:area-rectangle` with `svg:x`/`svg:y`/`svg:width`/`svg:height`.
    Rectangle {
        /// Left edge coordinate.
        x: String,
        /// Top edge coordinate.
        y: String,
        /// Area width.
        width: String,
        /// Area height.
        height: String,
    },
    /// `draw:area-circle` with `svg:cx`/`svg:cy`/`svg:r`.
    Circle {
        /// Center x coordinate.
        cx: String,
        /// Center y coordinate.
        cy: String,
        /// Radius.
        r: String,
    },
    /// `draw:area-polygon` with extents, view box, and point list.
    Polygon {
        /// Left edge coordinate of the extent.
        x: String,
        /// Top edge coordinate of the extent.
        y: String,
        /// Extent width.
        width: String,
        /// Extent height.
        height: String,
        /// `svg:viewBox` of the polygon coordinate space.
        view_box: String,
        /// `svg:points` vertex list.
        points: String,
    },
}

/// One clickable area of an image map, with inert link metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImageMapArea {
    /// The area geometry.
    pub shape: ImageMapAreaShape,
    /// `xlink:href` link target, stored verbatim and never resolved.
    pub href: Option<String>,
    /// `office:target-frame-name` link frame target.
    pub target_frame_name: Option<String>,
    /// `xlink:show` presentation hint (`new` or `replace`).
    pub show: Option<String>,
    /// `draw:nohref`: the area has no link target.
    pub no_href: bool,
    /// `office:name` of the area.
    pub name: Option<String>,
    /// Exact `svg:title` child XML, when present.
    pub title_xml: Option<String>,
    /// Exact `svg:desc` child XML, when present.
    pub description_xml: Option<String>,
    /// Exact `office:event-listeners` child XML, preserved without
    /// interpretation.
    pub event_listeners_xml: Option<String>,
    /// Exact area element XML.
    pub xml: String,
}

/// A `draw:image-map` element and its clickable areas.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImageMap {
    /// The clickable areas, in document order.
    pub areas: Vec<ImageMapArea>,
    /// Exact `draw:image-map` element XML.
    pub xml: String,
}

#[allow(clippy::type_complexity)]
fn parse_link_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<(
    Option<String>,
    Option<String>,
    Option<String>,
    bool,
    Option<String>,
)> {
    if let Some(link_type) = namespaced_attr(reader, element, XLINK, b"type")?
        && link_type != "simple"
    {
        return Err(invalid("image-map xlink:type must be 'simple'"));
    }
    let show = match namespaced_attr(reader, element, XLINK, b"show")?.as_deref() {
        None => None,
        Some(value @ ("new" | "replace")) => Some(value.to_string()),
        Some(value) => {
            return Err(invalid(format!(
                "image-map xlink:show must be 'new' or 'replace', found '{value}'"
            )));
        },
    };
    Ok((
        namespaced_attr(reader, element, XLINK, b"href")?,
        namespaced_attr(reader, element, OFFICE, b"target-frame-name")?,
        show,
        bool_attr(reader, element, DRAW, b"nohref")?,
        namespaced_attr(reader, element, OFFICE, b"name")?,
    ))
}

struct AreaCapture {
    area: ImageMapArea,
    start: usize,
    depth: usize,
    title_start: Option<usize>,
    description_start: Option<usize>,
    listeners_start: Option<usize>,
}

/// Parse every `draw:image-map` element in a document, in document order.
pub fn parse_image_maps(xml: &str) -> Result<Vec<ImageMap>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("image-map input exceeds 64 MiB"));
    }
    // quick-xml strips a UTF-8 BOM and reports positions relative to the
    // stripped text, so slice against the same view.
    let xml = xml.strip_prefix('\u{FEFF}').unwrap_or(xml);
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut maps = Vec::new();
    let mut active_map: Option<(Vec<ImageMapArea>, usize, usize)> = None;
    let mut active_area: Option<AreaCapture> = None;

    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("image-map parsing error: {error}")))?;
        let draw_element =
            matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == DRAW);
        let svg_element =
            matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == SVG);
        let office_element =
            matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == OFFICE);
        let event = event.into_owned();
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                if draw_element && element.local_name().as_ref() == b"image-map" {
                    if active_map.is_some() {
                        return Err(invalid("nested draw:image-map element"));
                    }
                    active_map = Some((Vec::new(), event_start, 1));
                } else if let Some((_, _, map_depth)) = active_map.as_mut() {
                    *map_depth += 1;
                    let local_name = element.local_name();
                    let local = local_name.as_ref();
                    let shape = if draw_element && local == b"area-rectangle" {
                        Some(ImageMapAreaShape::Rectangle {
                            x: required_attr(&reader, &element, SVG, b"x")?,
                            y: required_attr(&reader, &element, SVG, b"y")?,
                            width: required_attr(&reader, &element, SVG, b"width")?,
                            height: required_attr(&reader, &element, SVG, b"height")?,
                        })
                    } else if draw_element && local == b"area-circle" {
                        Some(ImageMapAreaShape::Circle {
                            cx: required_attr(&reader, &element, SVG, b"cx")?,
                            cy: required_attr(&reader, &element, SVG, b"cy")?,
                            r: required_attr(&reader, &element, SVG, b"r")?,
                        })
                    } else if draw_element && local == b"area-polygon" {
                        Some(ImageMapAreaShape::Polygon {
                            x: required_attr(&reader, &element, SVG, b"x")?,
                            y: required_attr(&reader, &element, SVG, b"y")?,
                            width: required_attr(&reader, &element, SVG, b"width")?,
                            height: required_attr(&reader, &element, SVG, b"height")?,
                            view_box: required_attr(&reader, &element, SVG, b"viewBox")?,
                            points: required_attr(&reader, &element, SVG, b"points")?,
                        })
                    } else {
                        None
                    };
                    if let Some(shape) = shape {
                        if active_area.is_some() || *map_depth != 2 {
                            return Err(invalid(
                                "image-map area must be a direct child of draw:image-map",
                            ));
                        }
                        let (href, target_frame_name, show, no_href, name) =
                            parse_link_attributes(&reader, &element)?;
                        active_area = Some(AreaCapture {
                            area: ImageMapArea {
                                shape,
                                href,
                                target_frame_name,
                                show,
                                no_href,
                                name,
                                title_xml: None,
                                description_xml: None,
                                event_listeners_xml: None,
                                xml: String::new(),
                            },
                            start: event_start,
                            depth: 1,
                            title_start: None,
                            description_start: None,
                            listeners_start: None,
                        });
                    } else if let Some(area) = active_area.as_mut() {
                        area.depth += 1;
                        if area.depth == 2 {
                            if svg_element && local == b"title" {
                                area.title_start = Some(event_start);
                            } else if svg_element && local == b"desc" {
                                area.description_start = Some(event_start);
                            } else if office_element && local == b"event-listeners" {
                                area.listeners_start = Some(event_start);
                            } else {
                                return Err(invalid(
                                    "image-map area contains an unexpected child element",
                                ));
                            }
                        }
                    } else {
                        return Err(invalid(format!(
                            "draw:image-map contains an unexpected {} child element",
                            String::from_utf8_lossy(local)
                        )));
                    }
                }
            },
            Event::Empty(element) => {
                let local_name = element.local_name();
                let local = local_name.as_ref();
                if let Some(area) = active_area.as_mut() {
                    if area.depth == 1 {
                        // Empty title/desc/event-listeners children are legal.
                        if svg_element && local == b"title" {
                            area.area.title_xml = Some(xml[event_start..event_end].to_string());
                        } else if svg_element && local == b"desc" {
                            area.area.description_xml =
                                Some(xml[event_start..event_end].to_string());
                        } else if office_element && local == b"event-listeners" {
                            area.area.event_listeners_xml =
                                Some(xml[event_start..event_end].to_string());
                        } else {
                            return Err(invalid(
                                "image-map area contains an unexpected child element",
                            ));
                        }
                    }
                } else if let Some((areas, _, _)) = active_map.as_mut() {
                    let shape = if draw_element && local == b"area-rectangle" {
                        Some(ImageMapAreaShape::Rectangle {
                            x: required_attr(&reader, &element, SVG, b"x")?,
                            y: required_attr(&reader, &element, SVG, b"y")?,
                            width: required_attr(&reader, &element, SVG, b"width")?,
                            height: required_attr(&reader, &element, SVG, b"height")?,
                        })
                    } else if draw_element && local == b"area-circle" {
                        Some(ImageMapAreaShape::Circle {
                            cx: required_attr(&reader, &element, SVG, b"cx")?,
                            cy: required_attr(&reader, &element, SVG, b"cy")?,
                            r: required_attr(&reader, &element, SVG, b"r")?,
                        })
                    } else if draw_element && local == b"area-polygon" {
                        Some(ImageMapAreaShape::Polygon {
                            x: required_attr(&reader, &element, SVG, b"x")?,
                            y: required_attr(&reader, &element, SVG, b"y")?,
                            width: required_attr(&reader, &element, SVG, b"width")?,
                            height: required_attr(&reader, &element, SVG, b"height")?,
                            view_box: required_attr(&reader, &element, SVG, b"viewBox")?,
                            points: required_attr(&reader, &element, SVG, b"points")?,
                        })
                    } else {
                        None
                    };
                    if let Some(shape) = shape {
                        let (href, target_frame_name, show, no_href, name) =
                            parse_link_attributes(&reader, &element)?;
                        areas.push(ImageMapArea {
                            shape,
                            href,
                            target_frame_name,
                            show,
                            no_href,
                            name,
                            title_xml: None,
                            description_xml: None,
                            event_listeners_xml: None,
                            xml: xml[event_start..event_end].to_string(),
                        });
                    } else {
                        return Err(invalid(format!(
                            "draw:image-map contains an unexpected {} child element",
                            String::from_utf8_lossy(local)
                        )));
                    }
                }
            },
            Event::End(element) => {
                let local_name = element.local_name();
                let local = local_name.as_ref();
                if let Some(area) = active_area.as_mut() {
                    if area.depth == 2 {
                        if area.title_start.is_some() && svg_element && local == b"title" {
                            let start = area.title_start.take().expect("title start");
                            area.area.title_xml = Some(xml[start..event_end].to_string());
                        } else if area.description_start.is_some()
                            && svg_element
                            && local == b"desc"
                        {
                            let start = area.description_start.take().expect("desc start");
                            area.area.description_xml = Some(xml[start..event_end].to_string());
                        } else if area.listeners_start.is_some()
                            && office_element
                            && local == b"event-listeners"
                        {
                            let start = area.listeners_start.take().expect("listeners start");
                            area.area.event_listeners_xml = Some(xml[start..event_end].to_string());
                        }
                    }
                    area.depth = area
                        .depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("image-map area nesting underflow"))?;
                    if area.depth == 0 {
                        let capture = active_area.take().expect("active area");
                        let mut area = capture.area;
                        area.xml = xml[capture.start..event_end].to_string();
                        let (areas, _, _) = active_map.as_mut().expect("active map");
                        areas.push(area);
                    }
                }
                if let Some((areas, map_start, map_depth)) = active_map.as_mut() {
                    *map_depth = map_depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("image-map nesting underflow"))?;
                    if *map_depth == 0 {
                        if !(draw_element && local == b"image-map") {
                            return Err(invalid("malformed draw:image-map element"));
                        }
                        let fragment = xml[*map_start..event_end].to_string();
                        maps.push(ImageMap {
                            areas: std::mem::take(areas),
                            xml: fragment,
                        });
                        active_map = None;
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if active_map.is_some() || active_area.is_some() {
        return Err(invalid("unterminated draw:image-map element"));
    }
    Ok(maps)
}

impl crate::OpenDocumentPackage {
    /// Every `draw:image-map` in packaged `content.xml`, in document order.
    pub fn image_maps(&self) -> Result<Vec<ImageMap>> {
        parse_image_maps(&self.content_xml()?)
    }
}

impl crate::FlatOpenDocument {
    /// Every `draw:image-map` in a flat OpenDocument, in document order.
    pub fn image_maps(&self) -> Result<Vec<ImageMap>> {
        parse_image_maps(self.xml())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const DOCUMENT: &str = concat!(
        r#"<?xml version="1.0"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
        r#"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" "#,
        r#"xmlns:xlink="http://www.w3.org/1999/xlink" "#,
        r#"office:version="1.3"><office:body><office:text>"#,
        r#"<draw:frame draw:name="Map">"#,
        r#"<draw:image-map>"#,
        r#"<draw:area-rectangle svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm" xlink:href="https://example.org/a" xlink:type="simple" office:target-frame-name="_blank" office:name="r1"><svg:title>Area A</svg:title></draw:area-rectangle>"#,
        r#"<draw:area-circle svg:cx="5cm" svg:cy="6cm" svg:r="7cm" draw:nohref="true"/>"#,
        r#"<draw:area-polygon svg:x="0cm" svg:y="0cm" svg:width="9cm" svg:height="8cm" svg:viewBox="0 0 100 100" svg:points="10,10 90,10 50,90" xlink:show="new" xlink:href="https://example.org/c"/>"#,
        r#"</draw:image-map>"#,
        r#"</draw:frame></office:text></office:body></office:document-content>"#,
    );

    #[test]
    fn parses_all_area_kinds_with_inert_links() {
        let maps = parse_image_maps(DOCUMENT).unwrap();
        assert_eq!(maps.len(), 1);
        let map = &maps[0];
        assert_eq!(map.areas.len(), 3);
        assert!(map.xml.starts_with("<draw:image-map>"));

        let ImageMapAreaShape::Rectangle { x, width, .. } = &map.areas[0].shape else {
            panic!()
        };
        assert_eq!(x, "1cm");
        assert_eq!(width, "3cm");
        assert_eq!(map.areas[0].href.as_deref(), Some("https://example.org/a"));
        assert_eq!(map.areas[0].target_frame_name.as_deref(), Some("_blank"));
        assert_eq!(map.areas[0].name.as_deref(), Some("r1"));
        assert!(!map.areas[0].no_href);
        assert_eq!(
            map.areas[0].title_xml.as_deref(),
            Some("<svg:title>Area A</svg:title>")
        );

        let ImageMapAreaShape::Circle { cx, r, .. } = &map.areas[1].shape else {
            panic!()
        };
        assert_eq!((cx.as_str(), r.as_str()), ("5cm", "7cm"));
        assert!(map.areas[1].no_href);
        assert!(map.areas[1].href.is_none());

        let ImageMapAreaShape::Polygon {
            view_box, points, ..
        } = &map.areas[2].shape
        else {
            panic!()
        };
        assert_eq!(view_box, "0 0 100 100");
        assert_eq!(points, "10,10 90,10 50,90");
        assert_eq!(map.areas[2].show.as_deref(), Some("new"));
    }

    #[test]
    fn reports_no_maps_when_absent() {
        let xml = concat!(
            r#"<?xml version="1.0"?><office:document-content "#,
            r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
            r#"office:version="1.3"><office:body/></office:document-content>"#,
        );
        assert!(parse_image_maps(xml).unwrap().is_empty());
    }

    #[test]
    fn rejects_malformed_maps() {
        // Missing required circle radius.
        let bad = DOCUMENT.replace(" svg:r=\"7cm\"", "");
        assert!(parse_image_maps(&bad).is_err());
        // Invalid xlink:show.
        let bad = DOCUMENT.replace("xlink:show=\"new\"", "xlink:show=\"embed\"");
        assert!(parse_image_maps(&bad).is_err());
        // Invalid xlink:type.
        let bad = DOCUMENT.replace("xlink:type=\"simple\"", "xlink:type=\"extended\"");
        assert!(parse_image_maps(&bad).is_err());
        // Invalid Boolean.
        let bad = DOCUMENT.replace("draw:nohref=\"true\"", "draw:nohref=\"maybe\"");
        assert!(parse_image_maps(&bad).is_err());
        // Unexpected child element.
        let bad = DOCUMENT.replace(
            "<svg:title>Area A</svg:title>",
            "<draw:glue-point draw:id=\"0\" svg:x=\"1cm\" svg:y=\"1cm\"/>",
        );
        assert!(parse_image_maps(&bad).is_err());
        // Nested image map.
        let bad = DOCUMENT.replace("<draw:area-circle", "<draw:image-map><draw:area-circle");
        assert!(parse_image_maps(&bad).is_err());
    }
}
