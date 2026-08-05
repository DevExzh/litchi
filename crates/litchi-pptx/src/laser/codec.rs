//! Bounded PresentationML laser-trace XML codec.
//!
//! Laser traces are retained as persisted presentation data only. This module
//! never replays, renders, interpolates, modifies, or executes slide-show
//! events.

use super::model::*;
use crate::time::{Offset, ParseError as TimeParseError};
use crate::{Error, Result};
use litchi_drawingml::coordinate::{Coordinate, ParseError as CoordinateParseError};
use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_ooxml_common::mce::{Capabilities, process_markup_compatibility};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;
use std::fmt::Write as _;

/// The PowerPoint extension URI that contains persisted laser-pointer traces.
pub const LASER_TRACE_EXTENSION_URI: &str = "{3A86A75C-4F4B-4683-9AE1-C65F6400EC91}";

const PRESENTATIONML_NAMESPACE_BYTES: &[u8] =
    b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PRESENTATIONML_NAMESPACE_BYTES: &[u8] =
    b"http://purl.oclc.org/ooxml/presentationml/main";

const P14_NAMESPACE: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const P14_NAMESPACE_BYTES: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2010/main";
const MAX_SLIDE_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_TOTAL_SLIDE_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_LASER_TRACES: usize = 4_096;
const MAX_LASER_POINTS: usize = 65_536;
const MAX_XML_NODES: usize = 250_000;
const MAX_XML_DEPTH: usize = 128;

#[derive(Clone, Copy, PartialEq, Eq)]
enum ElementKind {
    Other,
    Root,
    LaserExtension,
    LaserTraceList,
    LaserTrace,
    LaserPoint,
}

impl ElementKind {
    fn is_known(self) -> bool {
        matches!(
            self,
            Self::LaserExtension | Self::LaserTraceList | Self::LaserTrace | Self::LaserPoint
        )
    }
}

/// Read bounded, inert laser-pointer traces from one PresentationML slide.
pub fn read(slide_index: usize, xml_bytes: &[u8]) -> Result<Vec<Trace>> {
    read_with(slide_index, xml_bytes, &mut Limits::default())
}

/// Read one slide while accumulating resource use in limits.
pub fn read_with(slide_index: usize, xml_bytes: &[u8], limits: &mut Limits) -> Result<Vec<Trace>> {
    limits.add_slide_xml(xml_bytes.len())?;
    scan_slide_laser_traces(slide_index, xml_bytes, limits)
}

impl Limits {
    fn add_slide_xml(&mut self, bytes: usize) -> Result<()> {
        if bytes > MAX_SLIDE_XML_BYTES {
            return Err(limit("slide XML bytes", MAX_SLIDE_XML_BYTES));
        }
        self.total_slide_xml_bytes = self
            .total_slide_xml_bytes
            .checked_add(bytes)
            .ok_or_else(|| limit("total slide XML bytes", MAX_TOTAL_SLIDE_XML_BYTES))?;
        if self.total_slide_xml_bytes > MAX_TOTAL_SLIDE_XML_BYTES {
            return Err(limit("total slide XML bytes", MAX_TOTAL_SLIDE_XML_BYTES));
        }
        Ok(())
    }

    fn add_trace(&mut self) -> Result<()> {
        self.trace_count = self
            .trace_count
            .checked_add(1)
            .ok_or_else(|| limit("laser trace count", MAX_LASER_TRACES))?;
        if self.trace_count > MAX_LASER_TRACES {
            return Err(limit("laser trace count", MAX_LASER_TRACES));
        }
        Ok(())
    }

    fn add_point(&mut self) -> Result<()> {
        self.point_count = self
            .point_count
            .checked_add(1)
            .ok_or_else(|| limit("laser trace-point count", MAX_LASER_POINTS))?;
        if self.point_count > MAX_LASER_POINTS {
            return Err(limit("laser trace-point count", MAX_LASER_POINTS));
        }
        Ok(())
    }
}

fn scan_slide_laser_traces(
    slide_index: usize,
    xml_bytes: &[u8],
    limits: &mut Limits,
) -> Result<Vec<Trace>> {
    if xml_bytes.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("slide XML bytes", MAX_SLIDE_XML_BYTES));
    }

    let mut capabilities = Capabilities::ooxml_baseline();
    capabilities.understand_namespace(P14_NAMESPACE);
    let mce_limits = litchi_ooxml_common::mce::Limits {
        max_input_bytes: MAX_SLIDE_XML_BYTES,
        max_output_bytes: MAX_SLIDE_XML_BYTES,
        max_depth: MAX_XML_DEPTH,
        max_namespace_bindings: 4_096,
        max_directive_tokens: 4_096,
        max_choices_per_alternate: 1_024,
    };
    let xml = process_markup_compatibility(xml_bytes, &capabilities, &mce_limits)?.xml;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut stack = Vec::new();
    let mut traces = Vec::new();
    let mut active_trace = None;
    let mut nodes = 0usize;
    let mut saw_root = false;
    let mut closed_root = false;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                let depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| limit("slide XML depth", MAX_XML_DEPTH))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("slide XML depth", MAX_XML_DEPTH));
                }
                let parent = stack.last().copied().unwrap_or(ElementKind::Other);
                let kind = classify_element(
                    &namespace,
                    &element,
                    decoder,
                    parent,
                    depth,
                    saw_root,
                    false,
                    &mut active_trace,
                    limits,
                )?;
                if kind == ElementKind::Root {
                    saw_root = true;
                }
                stack.push(kind);
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                let depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| limit("slide XML depth", MAX_XML_DEPTH))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("slide XML depth", MAX_XML_DEPTH));
                }
                let parent = stack.last().copied().unwrap_or(ElementKind::Other);
                let kind = classify_element(
                    &namespace,
                    &element,
                    decoder,
                    parent,
                    depth,
                    saw_root,
                    true,
                    &mut active_trace,
                    limits,
                )?;
                if kind == ElementKind::Root {
                    saw_root = true;
                    closed_root = true;
                } else if kind == ElementKind::LaserTrace {
                    finish_trace(slide_index, &mut traces, &mut active_trace)?;
                }
            },
            Event::End(element) => {
                let kind = stack
                    .pop()
                    .ok_or_else(|| invalid("invalid slide XML nesting"))?;
                finish_element(
                    kind,
                    &namespace,
                    element.name(),
                    slide_index,
                    &mut traces,
                    &mut active_trace,
                )?;
                if kind == ElementKind::Root {
                    closed_root = true;
                }
            },
            Event::Text(text)
                if stack.last().copied().is_some_and(ElementKind::is_known)
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid("laser-trace markup cannot contain text"));
            },
            Event::CData(text)
                if stack.last().copied().is_some_and(ElementKind::is_known)
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid("laser-trace markup cannot contain text"));
            },
            Event::GeneralRef(_) if stack.last().copied().is_some_and(ElementKind::is_known) => {
                return Err(invalid(
                    "laser-trace markup cannot contain entity references",
                ));
            },
            Event::DocType(_) => return Err(invalid("slide XML must not contain a DTD")),
            Event::PI(_) => {
                return Err(invalid(
                    "slide XML must not contain a processing instruction",
                ));
            },
            Event::Eof => {
                if !stack.is_empty() || !saw_root || !closed_root {
                    return Err(invalid("unterminated or missing PresentationML slide root"));
                }
                break;
            },
            _ => {},
        }
    }

    Ok(traces)
}

#[allow(clippy::too_many_arguments)]
fn classify_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    parent: ElementKind,
    depth: usize,
    root_seen: bool,
    empty: bool,
    active_trace: &mut Option<Vec<TracePoint>>,
    limits: &mut Limits,
) -> Result<ElementKind> {
    if depth == 1 {
        if root_seen || !is_presentationml_name(namespace, element.name(), b"sld") {
            return Err(invalid(
                "slide XML must have one PresentationML sld root element",
            ));
        }
        return Ok(ElementKind::Root);
    }

    if is_presentationml_name(namespace, element.name(), b"ext")
        && is_laser_extension(element, decoder)?
    {
        return Ok(ElementKind::LaserExtension);
    }

    if is_p14_name(namespace, element.name(), b"laserTraceLst") {
        return match parent {
            ElementKind::LaserExtension => Ok(ElementKind::LaserTraceList),
            ElementKind::Other => Ok(ElementKind::Other),
            ElementKind::Root
            | ElementKind::LaserTraceList
            | ElementKind::LaserTrace
            | ElementKind::LaserPoint => Err(invalid(
                "laserTraceLst must be the direct child of its PowerPoint extension",
            )),
        };
    }

    if is_p14_name(namespace, element.name(), b"tracePtLst") {
        let kind = match parent {
            ElementKind::LaserTraceList => ElementKind::LaserTrace,
            ElementKind::Other => ElementKind::Other,
            ElementKind::Root
            | ElementKind::LaserExtension
            | ElementKind::LaserTrace
            | ElementKind::LaserPoint => {
                return Err(invalid(
                    "laser trace list contains tracePtLst in an invalid position",
                ));
            },
        };
        if kind == ElementKind::LaserTrace {
            limits.add_trace()?;
            if active_trace.replace(Vec::new()).is_some() {
                return Err(invalid("nested laser traces are not valid"));
            }
        }
        return Ok(kind);
    }

    if is_p14_name(namespace, element.name(), b"tracePt") {
        if parent != ElementKind::LaserTrace {
            return if parent.is_known() {
                Err(invalid(
                    "laser trace point is outside a PowerPoint laser trace",
                ))
            } else {
                Ok(ElementKind::Other)
            };
        }
        let point = parse_trace_point(element, decoder)?;
        limits.add_point()?;
        active_trace
            .as_mut()
            .ok_or_else(|| invalid("laser trace point has no active trace"))?
            .push(point);
        return Ok(if empty {
            ElementKind::Other
        } else {
            ElementKind::LaserPoint
        });
    }

    if parent.is_known() {
        return Err(invalid(
            "laser-trace extension contains an unsupported child element",
        ));
    }
    Ok(ElementKind::Other)
}

fn finish_element(
    kind: ElementKind,
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    slide_index: usize,
    traces: &mut Vec<Trace>,
    active_trace: &mut Option<Vec<TracePoint>>,
) -> Result<()> {
    match kind {
        ElementKind::Root if !is_presentationml_name(namespace, name, b"sld") => Err(invalid(
            "slide XML must close with a PresentationML sld element",
        )),
        ElementKind::LaserExtension if !is_presentationml_name(namespace, name, b"ext") => {
            Err(invalid("invalid laser-trace extension nesting"))
        },
        ElementKind::LaserTraceList if !is_p14_name(namespace, name, b"laserTraceLst") => {
            Err(invalid("invalid laser-trace-list nesting"))
        },
        ElementKind::LaserTrace if !is_p14_name(namespace, name, b"tracePtLst") => {
            Err(invalid("invalid laser-trace nesting"))
        },
        ElementKind::LaserTrace => finish_trace(slide_index, traces, active_trace),
        ElementKind::LaserPoint if !is_p14_name(namespace, name, b"tracePt") => {
            Err(invalid("invalid laser trace-point nesting"))
        },
        _ => Ok(()),
    }
}

fn finish_trace(
    slide_index: usize,
    traces: &mut Vec<Trace>,
    active_trace: &mut Option<Vec<TracePoint>>,
) -> Result<()> {
    let points = active_trace
        .take()
        .ok_or_else(|| invalid("laser trace has no active point list"))?;
    traces.push(Trace {
        slide_index,
        trace_index: traces.len(),
        points,
    });
    Ok(())
}

fn is_laser_extension(element: &BytesStart<'_>, decoder: Decoder) -> Result<bool> {
    Ok(
        unqualified_attribute_value(element, b"uri", decoder)?.as_deref()
            == Some(LASER_TRACE_EXTENSION_URI),
    )
}

fn is_presentationml_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
) -> bool {
    if name.local_name().as_ref() != local_name {
        return false;
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == PRESENTATIONML_NAMESPACE_BYTES
                || *value == STRICT_PRESENTATIONML_NAMESPACE_BYTES
        },
        ResolveResult::Unknown(prefix) => prefix.as_slice() == b"p",
        ResolveResult::Unbound => false,
    }
}

fn is_p14_name(namespace: &ResolveResult<'_>, name: QName<'_>, local_name: &[u8]) -> bool {
    name.local_name().as_ref() == local_name
        && matches!(
            namespace,
            ResolveResult::Bound(Namespace(value)) if *value == P14_NAMESPACE_BYTES
        )
}

fn parse_trace_point(element: &BytesStart<'_>, decoder: Decoder) -> Result<TracePoint> {
    let time = parse_time_offset(required_attribute(element, b"t", decoder)?)?;
    let x = parse_coordinate(required_attribute(element, b"x", decoder)?, "x")?;
    let y = parse_coordinate(required_attribute(element, b"y", decoder)?, "y")?;
    Ok(TracePoint { time, x, y })
}

fn required_attribute(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<String> {
    unqualified_attribute_value(element, name, decoder)?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            invalid(format!(
                "laser trace point is missing required '{}' attribute",
                String::from_utf8_lossy(name)
            ))
        })
}

fn parse_coordinate(value: String, name: &str) -> Result<Coordinate> {
    Coordinate::try_from(value).map_err(|error| coordinate_error(error, name))
}

fn coordinate_error(error: CoordinateParseError, name: &str) -> Error {
    invalid(format!(
        "invalid laser trace point {name} DrawingML coordinate: {error}"
    ))
}

fn parse_time_offset(value: String) -> Result<Offset> {
    Offset::try_from(value).map_err(time_error)
}

fn time_error(error: TimeParseError) -> Error {
    invalid(format!(
        "invalid laser trace point universal time offset: {error}"
    ))
}

fn increment_nodes(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("slide XML node count", MAX_XML_NODES))?;
    if *nodes > MAX_XML_NODES {
        return Err(limit("slide XML node count", MAX_XML_NODES));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

/// Check the bounded authoring domain for a laser trace.
pub fn validate(points: &[TracePoint]) -> Result<()> {
    if points.is_empty() {
        return Err(invalid("laser trace requires at least one point"));
    }
    if points.len() > MAX_LASER_POINTS {
        return Err(limit("laser trace-point count", MAX_LASER_POINTS));
    }
    Ok(())
}

/// Serialize one laser-pointer trace extension fragment.
pub fn write(points: &[TracePoint], conformance: Conformance) -> Result<String> {
    let mut xml = String::new();
    write_to(points, conformance, &mut xml)?;
    Ok(xml)
}

/// Append one laser-pointer trace extension fragment to an existing buffer.
pub fn write_to(points: &[TracePoint], conformance: Conformance, xml: &mut String) -> Result<()> {
    validate(points)?;
    let capacity = points
        .len()
        .checked_mul(48)
        .and_then(|bytes| bytes.checked_add(256))
        .ok_or_else(|| limit("laser trace XML bytes", MAX_SLIDE_XML_BYTES))?;
    xml.try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "laser trace XML",
            source,
        })?;

    xml.push_str("<p:ext xmlns:p=\"");
    xml.push_str(conformance.namespace());
    xml.push_str("\" xmlns:p14=\"");
    xml.push_str(P14_NAMESPACE);
    xml.push_str("\" uri=\"");
    xml.push_str(LASER_TRACE_EXTENSION_URI);
    xml.push_str("\"><p14:laserTraceLst><p14:tracePtLst>");
    for point in points {
        xml.push_str("<p14:tracePt t=\"");
        xml.push_str(point.time.as_str());
        xml.push_str("\" x=\"");
        write!(xml, "{}", point.x).map_err(|_| Error::Write)?;
        xml.push_str("\" y=\"");
        write!(xml, "{}", point.y).map_err(|_| Error::Write)?;
        xml.push_str("\"/>");
    }
    xml.push_str("</p14:tracePtLst></p14:laserTraceLst></p:ext>");
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    #[test]
    fn scans_laser_traces_through_markup_compatibility() {
        let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:mc="{MCE}" xmlns:p14="{P14_NAMESPACE}" mc:Ignorable="p14"><p:extLst><p:ext uri="{LASER_TRACE_EXTENSION_URI}"><p14:laserTraceLst><p14:tracePtLst><p14:tracePt t="1.5s" x="-2" y="3"/><p14:tracePt t="2000ms" x="1.5cm" y="-2pt"/></p14:tracePtLst><p14:tracePtLst/></p14:laserTraceLst></p:ext></p:extLst></p:sld>"#
        );
        let traces = read_with(4, xml.as_bytes(), &mut Limits::default()).unwrap();

        assert_eq!(traces.len(), 2);
        assert_eq!(traces[0].slide_index(), 4);
        assert_eq!(traces[0].trace_index(), 0);
        assert_eq!(traces[0].point_count(), 2);
        assert_eq!(traces[0].points()[0].time(), &Offset::ms(1500));
        assert_eq!(traces[0].points()[0].x().as_emu(), Some(-2));
        assert_eq!(traces[0].points()[1].x().to_string(), "1.5cm");
        assert_eq!(traces[0].points()[1].y().to_string(), "-2pt");
        assert_eq!(traces[1].trace_index(), 1);
        assert!(traces[1].points().is_empty());
    }

    #[test]
    fn rejects_malformed_laser_trace_points() {
        let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:p14="{P14_NAMESPACE}"><p:extLst><p:ext uri="{LASER_TRACE_EXTENSION_URI}"><p14:laserTraceLst><p14:tracePtLst><p14:tracePt t="1..2s" x="0" y="0"/></p14:tracePtLst></p14:laserTraceLst></p:ext></p:extLst></p:sld>"#
        );

        assert!(read_with(0, xml.as_bytes(), &mut Limits::default()).is_err());

        for coordinate in ["27273042316901", "1e2mm", "+1mm", "1px"] {
            let xml = format!(
                r#"<p:sld xmlns:p="{PML}" xmlns:p14="{P14_NAMESPACE}"><p:extLst><p:ext uri="{LASER_TRACE_EXTENSION_URI}"><p14:laserTraceLst><p14:tracePtLst><p14:tracePt t="0" x="{coordinate}" y="0"/></p14:tracePtLst></p14:laserTraceLst></p:ext></p:extLst></p:sld>"#
            );
            assert!(
                read_with(0, xml.as_bytes(), &mut Limits::default()).is_err(),
                "accepted {coordinate:?}"
            );
        }
    }

    #[test]
    fn writer_preserves_the_selected_presentation_dialect() {
        let point = TracePoint::new(
            Offset::ZERO,
            Coordinate::emu(914_400).unwrap(),
            Coordinate::emu(457_200).unwrap(),
        );
        let transitional = write(std::slice::from_ref(&point), Conformance::Transitional).unwrap();
        let strict = write(&[point], Conformance::Strict).unwrap();

        assert!(transitional.contains(PRESENTATIONML_NAMESPACE));
        assert!(strict.contains(STRICT_PRESENTATIONML_NAMESPACE));
    }
}
