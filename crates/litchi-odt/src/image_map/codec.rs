use litchi_core::{Error, Result};
use quick_xml::{
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

use super::{ImageMap, ImageMapArea, ImageMapAreaShape};

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
                            let start = area
                                .title_start
                                .take()
                                .ok_or_else(|| invalid("missing image-map title start"))?;
                            area.area.title_xml = Some(xml[start..event_end].to_string());
                        } else if area.description_start.is_some()
                            && svg_element
                            && local == b"desc"
                        {
                            let start = area
                                .description_start
                                .take()
                                .ok_or_else(|| invalid("missing image-map description start"))?;
                            area.area.description_xml = Some(xml[start..event_end].to_string());
                        } else if area.listeners_start.is_some()
                            && office_element
                            && local == b"event-listeners"
                        {
                            let start = area
                                .listeners_start
                                .take()
                                .ok_or_else(|| invalid("missing image-map listeners start"))?;
                            area.area.event_listeners_xml = Some(xml[start..event_end].to_string());
                        }
                    }
                    area.depth = area
                        .depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("image-map area nesting underflow"))?;
                    if area.depth == 0 {
                        let capture = active_area
                            .take()
                            .ok_or_else(|| invalid("missing completed image-map area"))?;
                        let mut area = capture.area;
                        area.xml = xml[capture.start..event_end].to_string();
                        let (areas, _, _) = active_map
                            .as_mut()
                            .ok_or_else(|| invalid("image-map area has no parent map"))?;
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
