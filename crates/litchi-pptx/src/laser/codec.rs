//! Bounded `PresentationML` laser-trace XML codec.
//!
//! Laser traces are retained as persisted presentation data only. This module
//! never replays, renders, interpolates, modifies, or executes slide-show
//! events.

use super::model::{Conformance, Limits, Trace, TracePoint, Usage};
use crate::time::{Offset, ParseError as TimeParseError};
use crate::{Error, Result};
use litchi_drawingml::coordinate::{Coordinate, ParseError as CoordinateParseError};
use litchi_ooxml_common::mce::{Capabilities, process_markup_compatibility};
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

/// The `PowerPoint` extension URI that contains persisted laser-pointer traces.
pub const LASER_TRACE_EXTENSION_URI: &str = "{3A86A75C-4F4B-4683-9AE1-C65F6400EC91}";

const PRESENTATIONML_NAMESPACE_BYTES: &[u8] =
    b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PRESENTATIONML_NAMESPACE_BYTES: &[u8] =
    b"http://purl.oclc.org/ooxml/presentationml/main";

const P14_NAMESPACE: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const P14_NAMESPACE_BYTES: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2010/main";
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

/// Read one slide under the caller-selected policy and accumulate successful
/// resource use in `usage`. The separate ledger is required so a package can
/// retain aggregate usage across all of its slide parts without mutating its
/// policy.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn read(
    slide_index: usize,
    xml_bytes: &[u8],
    limits: &Limits,
    usage: &mut Usage,
) -> Result<Vec<Trace>> {
    let mut next = *usage;
    limits.add_slide_xml(xml_bytes.len(), &mut next)?;
    let traces = scan_slide_laser_traces(slide_index, xml_bytes, limits, &mut next)?;
    *usage = next;
    Ok(traces)
}

impl Limits {
    fn add_slide_xml(&self, bytes: usize, usage: &mut Usage) -> Result<()> {
        if bytes > self.max_slide_xml_bytes() {
            return Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                bytes,
                self.max_slide_xml_bytes(),
            ));
        }
        let total = usage.add_slide_xml_bytes(bytes).ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::XmlBytes,
                usize::MAX,
                self.max_total_slide_xml_bytes(),
            )
        })?;
        if total > self.max_total_slide_xml_bytes() {
            return Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                total,
                self.max_total_slide_xml_bytes(),
            ));
        }
        Ok(())
    }

    fn add_trace(&self, usage: &mut Usage) -> Result<()> {
        let total = usage.add_trace().ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::Traces,
                usize::MAX,
                self.max_trace_count(),
            )
        })?;
        if total > self.max_trace_count() {
            return Err(limit(
                crate::error::CodecLimitKind::Traces,
                total,
                self.max_trace_count(),
            ));
        }
        Ok(())
    }

    fn add_point(&self, usage: &mut Usage) -> Result<()> {
        let total = usage.add_point().ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::TracePoints,
                usize::MAX,
                self.max_point_count(),
            )
        })?;
        if total > self.max_point_count() {
            return Err(limit(
                crate::error::CodecLimitKind::TracePoints,
                total,
                self.max_point_count(),
            ));
        }
        Ok(())
    }
}

fn scan_slide_laser_traces(
    slide_index: usize,
    xml_bytes: &[u8],
    limits: &Limits,
    usage: &mut Usage,
) -> Result<Vec<Trace>> {
    if xml_bytes.len() > limits.max_slide_xml_bytes() {
        return Err(limit(
            crate::error::CodecLimitKind::XmlBytes,
            xml_bytes.len(),
            limits.max_slide_xml_bytes(),
        ));
    }

    let mut capabilities = Capabilities::ooxml_baseline();
    capabilities.understand_namespace(P14_NAMESPACE);
    let mce_limits = litchi_ooxml_common::mce::Limits {
        max_input_bytes: limits.max_slide_xml_bytes(),
        max_output_bytes: limits.max_slide_xml_bytes(),
        max_depth: limits.max_xml_depth(),
        max_namespace_bindings: 4_096,
        max_directive_tokens: 4_096,
        max_choices_per_alternate: 1_024,
        max_attributes_per_element: litchi_ooxml_common::mce::DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT,
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
                increment_nodes(&mut nodes, limits)?;
                let depth = stack.len().checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        limits.max_xml_depth(),
                    )
                })?;
                if depth > limits.max_xml_depth() {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        depth,
                        limits.max_xml_depth(),
                    ));
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
                    usage,
                )?;
                if kind == ElementKind::Root {
                    saw_root = true;
                }
                stack.push(kind);
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes, limits)?;
                let depth = stack.len().checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        limits.max_xml_depth(),
                    )
                })?;
                if depth > limits.max_xml_depth() {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        depth,
                        limits.max_xml_depth(),
                    ));
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
                    usage,
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

#[allow(
    clippy::too_many_arguments,
    reason = "element classifier threads one slot per laser-element field"
)]
fn classify_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    parent: ElementKind,
    depth: usize,
    root_seen: bool,
    empty: bool,
    active_trace: &mut Option<Vec<TracePoint>>,
    limits: &Limits,
    usage: &mut Usage,
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
            limits.add_trace(usage)?;
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
        limits.add_point(usage)?;
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
    match error {
        CoordinateParseError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Drawing(litchi_drawingml::Error::Limit {
            kind,
            observed,
            maximum,
        }),
        _ => invalid(format!(
            "invalid laser trace point {name} DrawingML coordinate"
        )),
    }
}

fn parse_time_offset(value: String) -> Result<Offset> {
    Offset::try_from(value).map_err(time_error)
}

fn time_error(error: TimeParseError) -> Error {
    invalid(format!(
        "invalid laser trace point universal time offset: {error}"
    ))
}

fn increment_nodes(nodes: &mut usize, limits: &Limits) -> Result<()> {
    *nodes = nodes.checked_add(1).ok_or_else(|| {
        limit(
            crate::error::CodecLimitKind::XmlNodes,
            usize::MAX,
            limits.max_xml_nodes(),
        )
    })?;
    if *nodes > limits.max_xml_nodes() {
        return Err(limit(
            crate::error::CodecLimitKind::XmlNodes,
            *nodes,
            limits.max_xml_nodes(),
        ));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

/// # Errors
///
/// Returns an error if the operation fails.
fn limit(kind: crate::error::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: crate::error::LimitKind::Codec(kind),
        observed,
        maximum,
    }
}

/// Check the bounded authoring domain for a laser trace.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn validate(points: &[TracePoint], limits: &Limits) -> Result<()> {
    if points.is_empty() {
        return Err(invalid("laser trace requires at least one point"));
    }
    if points.len() > limits.max_point_count() {
        return Err(limit(
            crate::error::CodecLimitKind::TracePoints,
            points.len(),
            limits.max_point_count(),
        ));
    }
    Ok(())
}

/// Serialize one laser-pointer trace extension fragment.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn write(points: &[TracePoint], conformance: Conformance, limits: Limits) -> Result<String> {
    let mut xml = String::new();
    write_to(points, conformance, &mut xml, limits)?;
    Ok(xml)
}

/// Append one laser-pointer trace extension fragment to an existing buffer.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn write_to(
    points: &[TracePoint],
    conformance: Conformance,
    xml: &mut String,
    limits: Limits,
) -> Result<()> {
    validate(points, &limits)?;

    // Coordinates and timestamps are variable-width values.  Count the
    // complete fragment first so the destination gets one exact, fallible
    // reservation instead of a payload-dependent estimate.
    let mut count = CountingXml::default();
    emit(points, conformance, &mut count)?;
    let output_len = count.finish()?;
    if output_len > limits.max_output_bytes() {
        return Err(limit(
            crate::error::CodecLimitKind::OutputBytes,
            output_len,
            limits.max_output_bytes(),
        ));
    }

    let start = xml.len();
    xml.try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "laser trace XML",
            source,
        })?;

    let result = {
        let mut output = LimitedXml {
            xml,
            start,
            maximum: output_len,
        };
        let result = emit(points, conformance, &mut output);
        result.and_then(|()| {
            if output.xml.len().saturating_sub(start) != output_len {
                Err(invalid("laser trace XML length changed during emission"))
            } else {
                Ok(())
            }
        })
    };
    if result.is_err() {
        xml.truncate(start);
    }
    result
}

struct LimitedXml<'a> {
    xml: &'a mut String,
    start: usize,
    maximum: usize,
}

trait LaserSink {
    fn push(&mut self, value: &str) -> Result<()>;
    fn format(&mut self, arguments: std::fmt::Arguments<'_>) -> Result<()>;
}

#[derive(Default)]
struct CountingXml {
    length: usize,
    overflowed: bool,
}

impl CountingXml {
    fn finish(self) -> Result<usize> {
        if self.overflowed {
            Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                usize::MAX,
                usize::MAX,
            ))
        } else {
            Ok(self.length)
        }
    }
}

impl LaserSink for CountingXml {
    fn push(&mut self, value: &str) -> Result<()> {
        if let Some(length) = self.length.checked_add(value.len()) {
            self.length = length;
            Ok(())
        } else {
            self.overflowed = true;
            Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                usize::MAX,
                usize::MAX,
            ))
        }
    }

    fn format(&mut self, arguments: std::fmt::Arguments<'_>) -> Result<()> {
        std::fmt::write(self, arguments).map_err(|_| {
            limit(
                crate::error::CodecLimitKind::XmlBytes,
                usize::MAX,
                usize::MAX,
            )
        })
    }
}

impl std::fmt::Write for CountingXml {
    fn write_str(&mut self, value: &str) -> std::fmt::Result {
        self.push(value).map_err(|_| std::fmt::Error)
    }
}

impl LimitedXml<'_> {
    fn push(&mut self, value: &str) -> Result<()> {
        let Some(length) = self
            .xml
            .len()
            .checked_sub(self.start)
            .and_then(|length| length.checked_add(value.len()))
        else {
            return Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                usize::MAX,
                self.maximum,
            ));
        };
        let Some(end) = self.xml.len().checked_add(value.len()) else {
            return Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                usize::MAX,
                self.maximum,
            ));
        };
        if length > self.maximum || end > self.xml.capacity() {
            return Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                length.max(end),
                self.maximum,
            ));
        }
        self.xml.push_str(value);
        Ok(())
    }

    fn format(&mut self, arguments: std::fmt::Arguments<'_>) -> Result<()> {
        std::fmt::write(self, arguments).map_err(|_| {
            limit(
                crate::error::CodecLimitKind::XmlBytes,
                usize::MAX,
                self.maximum,
            )
        })
    }
}

impl LaserSink for LimitedXml<'_> {
    fn push(&mut self, value: &str) -> Result<()> {
        Self::push(self, value)
    }

    fn format(&mut self, arguments: std::fmt::Arguments<'_>) -> Result<()> {
        Self::format(self, arguments)
    }
}

fn emit<S: LaserSink>(
    points: &[TracePoint],
    conformance: Conformance,
    output: &mut S,
) -> Result<()> {
    output.push("<p:ext xmlns:p=\"")?;
    output.push(conformance.namespace())?;
    output.push("\" xmlns:p14=\"")?;
    output.push(P14_NAMESPACE)?;
    output.push("\" uri=\"")?;
    output.push(LASER_TRACE_EXTENSION_URI)?;
    output.push("\"><p14:laserTraceLst><p14:tracePtLst>")?;
    for point in points {
        output.push("<p14:tracePt t=\"")?;
        output.push(point.time.as_str())?;
        output.push("\" x=\"")?;
        output.format(format_args!("{}", point.x))?;
        output.push("\" y=\"")?;
        output.format(format_args!("{}", point.y))?;
        output.push("\"/>")?;
    }
    output.push("</p14:tracePtLst></p14:laserTraceLst></p:ext>")
}

impl std::fmt::Write for LimitedXml<'_> {
    fn write_str(&mut self, value: &str) -> std::fmt::Result {
        if self.push(value).is_err() {
            return Err(std::fmt::Error);
        }
        Ok(())
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    #[test]
    fn scans_laser_traces_through_markup_compatibility() {
        let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:mc="{MCE}" xmlns:p14="{P14_NAMESPACE}" mc:Ignorable="p14"><p:extLst><p:ext uri="{LASER_TRACE_EXTENSION_URI}"><p14:laserTraceLst><p14:tracePtLst><p14:tracePt t="1.5s" x="-2" y="3"/><p14:tracePt t="2000ms" x="1.5cm" y="-2pt"/></p14:tracePtLst><p14:tracePtLst/></p14:laserTraceLst></p:ext></p:extLst></p:sld>"#
        );
        let mut usage = Usage::default();
        let traces = read(4, xml.as_bytes(), &Limits::DEFAULT, &mut usage).unwrap();

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

        assert!(read(0, xml.as_bytes(), &Limits::DEFAULT, &mut Usage::default()).is_err());

        for coordinate in ["27273042316901", "1e2mm", "+1mm", "1px"] {
            let xml = format!(
                r#"<p:sld xmlns:p="{PML}" xmlns:p14="{P14_NAMESPACE}"><p:extLst><p:ext uri="{LASER_TRACE_EXTENSION_URI}"><p14:laserTraceLst><p14:tracePtLst><p14:tracePt t="0" x="{coordinate}" y="0"/></p14:tracePtLst></p14:laserTraceLst></p:ext></p:extLst></p:sld>"#
            );
            assert!(
                read(0, xml.as_bytes(), &Limits::DEFAULT, &mut Usage::default()).is_err(),
                "accepted {coordinate:?}"
            );
        }
    }

    #[test]
    fn usage_preserves_aggregate_across_slides_without_mutating_policy() {
        let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:p14="{P14_NAMESPACE}"><p:extLst><p:ext uri="{LASER_TRACE_EXTENSION_URI}"><p14:laserTraceLst><p14:tracePtLst><p14:tracePt t="0" x="0" y="0"/></p14:tracePtLst></p14:laserTraceLst></p:ext></p:extLst></p:sld>"#
        );
        let limits = Limits::new(
            xml.len(),
            xml.len().checked_mul(2).unwrap(),
            2,
            2,
            Limits::DEFAULT.max_xml_nodes(),
            Limits::DEFAULT.max_xml_depth(),
            Limits::DEFAULT.max_output_bytes(),
        )
        .unwrap();
        let mut usage = Usage::default();
        read(0, xml.as_bytes(), &limits, &mut usage).unwrap();
        read(1, xml.as_bytes(), &limits, &mut usage).unwrap();
        assert_eq!(usage.total_slide_xml_bytes(), xml.len() * 2);
        assert_eq!(usage.point_count(), 2);
        assert_eq!(limits.max_point_count(), 2);

        let before_failed_read = usage;
        assert!(read(2, xml.as_bytes(), &limits, &mut usage).is_err());
        assert_eq!(usage, before_failed_read);
    }

    #[test]
    fn writer_preserves_the_selected_presentation_dialect() {
        let point = TracePoint::new(
            Offset::ZERO,
            Coordinate::emu(914_400).unwrap(),
            Coordinate::emu(457_200).unwrap(),
        );
        let transitional = write(
            std::slice::from_ref(&point),
            Conformance::Transitional,
            Limits::DEFAULT,
        )
        .unwrap();
        let strict = write(&[point], Conformance::Strict, Limits::DEFAULT).unwrap();

        assert!(
            transitional.contains("http://schemas.openxmlformats.org/presentationml/2006/main")
        );
        assert!(strict.contains("http://purl.oclc.org/ooxml/presentationml/main"));
    }

    #[test]
    fn writer_admits_exact_output_and_refuses_one_byte_under_atomically() {
        let point = TracePoint::new(
            Offset::ZERO,
            Coordinate::emu(914_400).unwrap(),
            Coordinate::emu(457_200).unwrap(),
        );
        let expected = write(
            std::slice::from_ref(&point),
            Conformance::Transitional,
            Limits::DEFAULT,
        )
        .unwrap();
        let exact = Limits::DEFAULT
            .with_output_bytes(expected.len())
            .expect("bounded output policy");
        assert_eq!(
            write(
                std::slice::from_ref(&point),
                Conformance::Transitional,
                exact
            )
            .unwrap(),
            expected
        );

        let narrow = Limits::DEFAULT
            .with_output_bytes(expected.len() - 1)
            .expect("one-byte narrower output policy");
        let mut output = "prefix".to_owned();
        assert!(matches!(
            write_to(
                std::slice::from_ref(&point),
                Conformance::Transitional,
                &mut output,
                narrow,
            ),
            Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
                ..
            })
        ));
        assert_eq!(output, "prefix");
    }
}
