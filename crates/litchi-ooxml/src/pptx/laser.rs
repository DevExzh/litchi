//! Bounded, inert PowerPoint laser-pointer trace discovery.
//!
//! Laser traces are retained as persisted presentation data only. This module
//! never replays, renders, interpolates, modifies, or executes slide-show
//! events.

use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::is_presentationml_name;
use litchi_drawingml::coord::{Coordinate, ParseError as CoordinateParseError};
use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_ooxml_common::{MceCapabilities, MceLimits, process_markup_compatibility};
use litchi_opc::Part;
use litchi_opc::constants::content_type as ct;
use litchi_pptx::time::{Offset, ParseError as TimeParseError};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;
use std::fmt::Write as _;

/// The PowerPoint extension URI that contains persisted laser-pointer traces.
pub const LASER_TRACE_EXTENSION_URI: &str = "{3A86A75C-4F4B-4683-9AE1-C65F6400EC91}";

const P14_NAMESPACE: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const P14_NAMESPACE_BYTES: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2010/main";
const MAX_SLIDE_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_TOTAL_SLIDE_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_LASER_TRACES: usize = 4_096;
const MAX_LASER_POINTS: usize = 65_536;
const MAX_XML_NODES: usize = 250_000;
const MAX_XML_DEPTH: usize = 128;

/// A persisted laser-pointer point from a PowerPoint slide show.
///
/// The represented duration is exact; its source spelling is canonicalized to
/// a normalized typed offset.
/// Coordinates are exact `a:ST_Coordinate` values relative to the slide's
/// top-left corner.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PptxLaserTracePoint {
    time: Offset,
    x: Coordinate,
    y: Coordinate,
}

impl PptxLaserTracePoint {
    /// Return the exact normalized time offset relative to the slide timeline.
    #[inline]
    pub fn time(&self) -> &Offset {
        &self.time
    }

    /// Return the checked horizontal DrawingML coordinate.
    #[inline]
    pub fn x(&self) -> &Coordinate {
        &self.x
    }

    /// Return the checked vertical DrawingML coordinate.
    #[inline]
    pub fn y(&self) -> &Coordinate {
        &self.y
    }
}

/// A bounded, inert laser-pointer trace recorded for a presentation slide.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PptxLaserTrace {
    slide_index: usize,
    trace_index: usize,
    points: Vec<PptxLaserTracePoint>,
}

impl PptxLaserTrace {
    /// Return the zero-based index of the slide that owns this trace.
    #[inline]
    pub fn slide_index(&self) -> usize {
        self.slide_index
    }

    /// Return the zero-based source-order index of this trace on its slide.
    #[inline]
    pub fn trace_index(&self) -> usize {
        self.trace_index
    }

    /// Return the stored trace points in source order.
    #[inline]
    pub fn points(&self) -> &[PptxLaserTracePoint] {
        &self.points
    }

    /// Return the number of stored trace points.
    #[inline]
    pub fn point_count(&self) -> usize {
        self.points.len()
    }
}

#[derive(Default)]
pub(crate) struct LaserLoadLimits {
    total_slide_xml_bytes: usize,
    trace_count: usize,
    point_count: usize,
}

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

/// Load bounded, inert laser-pointer traces from one PresentationML slide.
pub(crate) fn load_slide_laser_traces(
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut LaserLoadLimits,
) -> Result<Vec<PptxLaserTrace>> {
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid(
            "laser-trace discovery requires a PresentationML slide part",
        ));
    }
    limits.add_slide_xml(slide.blob().len())?;
    scan_slide_laser_traces(slide_index, slide.blob(), limits)
}

impl LaserLoadLimits {
    fn add_slide_xml(&mut self, bytes: usize) -> Result<()> {
        if bytes > MAX_SLIDE_XML_BYTES {
            return Err(limit("slide XML bytes"));
        }
        self.total_slide_xml_bytes = self
            .total_slide_xml_bytes
            .checked_add(bytes)
            .ok_or_else(|| limit("total slide XML bytes"))?;
        if self.total_slide_xml_bytes > MAX_TOTAL_SLIDE_XML_BYTES {
            return Err(limit("total slide XML bytes"));
        }
        Ok(())
    }

    fn add_trace(&mut self) -> Result<()> {
        self.trace_count = self
            .trace_count
            .checked_add(1)
            .ok_or_else(|| limit("laser trace count"))?;
        if self.trace_count > MAX_LASER_TRACES {
            return Err(limit("laser trace count"));
        }
        Ok(())
    }

    fn add_point(&mut self) -> Result<()> {
        self.point_count = self
            .point_count
            .checked_add(1)
            .ok_or_else(|| limit("laser trace-point count"))?;
        if self.point_count > MAX_LASER_POINTS {
            return Err(limit("laser trace-point count"));
        }
        Ok(())
    }
}

fn scan_slide_laser_traces(
    slide_index: usize,
    xml_bytes: &[u8],
    limits: &mut LaserLoadLimits,
) -> Result<Vec<PptxLaserTrace>> {
    if xml_bytes.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("slide XML bytes"));
    }

    let mut capabilities = MceCapabilities::ooxml_baseline();
    capabilities.understand_namespace(P14_NAMESPACE);
    let mce_limits = MceLimits {
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
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                let depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| limit("slide XML depth"))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("slide XML depth"));
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
                    .ok_or_else(|| limit("slide XML depth"))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("slide XML depth"));
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
    active_trace: &mut Option<Vec<PptxLaserTracePoint>>,
    limits: &mut LaserLoadLimits,
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
    traces: &mut Vec<PptxLaserTrace>,
    active_trace: &mut Option<Vec<PptxLaserTracePoint>>,
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
    traces: &mut Vec<PptxLaserTrace>,
    active_trace: &mut Option<Vec<PptxLaserTracePoint>>,
) -> Result<()> {
    let points = active_trace
        .take()
        .ok_or_else(|| invalid("laser trace has no active point list"))?;
    traces.push(PptxLaserTrace {
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

fn is_p14_name(namespace: &ResolveResult<'_>, name: QName<'_>, local_name: &[u8]) -> bool {
    name.local_name().as_ref() == local_name
        && matches!(
            namespace,
            ResolveResult::Bound(Namespace(value)) if *value == P14_NAMESPACE_BYTES
        )
}

fn parse_trace_point(element: &BytesStart<'_>, decoder: Decoder) -> Result<PptxLaserTracePoint> {
    let time = parse_time_offset(required_attribute(element, b"t", decoder)?)?;
    let x = parse_coordinate(required_attribute(element, b"x", decoder)?, "x")?;
    let y = parse_coordinate(required_attribute(element, b"y", decoder)?, "y")?;
    Ok(PptxLaserTracePoint { time, x, y })
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

fn coordinate_error(error: CoordinateParseError, name: &str) -> OoxmlError {
    invalid(format!(
        "invalid laser trace point {name} DrawingML coordinate: {error}"
    ))
}

fn parse_time_offset(value: String) -> Result<Offset> {
    Offset::try_from(value).map_err(time_error)
}

fn time_error(error: TimeParseError) -> OoxmlError {
    invalid(format!(
        "invalid laser trace point universal time offset: {error}"
    ))
}

fn increment_nodes(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("slide XML node count"))?;
    if *nodes > MAX_XML_NODES {
        return Err(limit("slide XML node count"));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn limit(what: &str) -> OoxmlError {
    invalid(format!("{what} exceeds the supported safety limit"))
}

impl PptxLaserTracePoint {
    /// Create a trace point from exact, checked time and coordinate values.
    pub fn new(time: Offset, x: Coordinate, y: Coordinate) -> Self {
        Self { time, x, y }
    }
}

/// Store one laser-pointer trace onto a slide as a PowerPoint 2010
/// `p14:laserTraceLst` extension.
///
/// The points are typed and serialized canonically into a new
/// `p14:tracePtLst`; the slide gains the `p:ext` extension block (creating
/// `p:extLst` when absent) while preserving its namespace dialect. Slides
/// that already carry a laser extension are rejected — replacement is not
/// supported in this pass. Traces are never replayed, rendered,
/// interpolated, or executed.
pub fn store_slide_laser_trace(
    package: &mut litchi_opc::OpcPackage,
    slide_name: &litchi_opc::PackURI,
    points: &[PptxLaserTracePoint],
) -> Result<()> {
    if points.is_empty() {
        return Err(invalid("laser trace requires at least one point"));
    }
    if points.len() > MAX_LASER_POINTS {
        return Err(limit("laser trace-point count"));
    }
    let slide = package.get_part(slide_name)?;
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid(
            "laser-trace storage requires a PresentationML slide part",
        ));
    }
    if !load_slide_laser_traces(0, slide, &mut LaserLoadLimits::default())?.is_empty() {
        return Err(invalid(
            "slide already contains a laser-trace extension; replacement is not supported",
        ));
    }

    let mut fragment = String::with_capacity(points.len() * 48 + 256);
    fragment.push_str("<p:ext xmlns:p=\"");
    fragment.push_str(crate::pptx::slide_patch::slide_dialect(slide.blob())?);
    fragment.push_str("\" xmlns:p14=\"");
    fragment.push_str(P14_NAMESPACE);
    fragment.push_str("\" uri=\"");
    fragment.push_str(LASER_TRACE_EXTENSION_URI);
    fragment.push_str("\"><p14:laserTraceLst><p14:tracePtLst>");
    for point in points {
        fragment.push_str("<p14:tracePt t=\"");
        fragment.push_str(point.time.as_str());
        fragment.push_str("\" x=\"");
        write!(fragment, "{}", point.x)
            .map_err(|_| invalid("failed to serialize laser trace x coordinate"))?;
        fragment.push_str("\" y=\"");
        write!(fragment, "{}", point.y)
            .map_err(|_| invalid("failed to serialize laser trace y coordinate"))?;
        fragment.push_str("\"/>");
    }
    fragment.push_str("</p14:tracePtLst></p14:laserTraceLst></p:ext>");

    let updated = crate::pptx::slide_patch::insert_extension_fragment(slide.blob(), &fragment)?;
    // Self-check: the patched slide must read back through the discovery path.
    let probe =
        litchi_opc::BlobPart::new(slide_name.clone(), ct::PML_SLIDE.into(), updated.clone());
    let traces = load_slide_laser_traces(0, &probe, &mut LaserLoadLimits::default())?;
    if traces.len() != 1 || traces[0].points().len() != points.len() {
        return Err(invalid("laser-trace storage failed read-back validation"));
    }
    package.get_part_mut(slide_name)?.set_blob(updated);
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
        let traces =
            scan_slide_laser_traces(4, xml.as_bytes(), &mut LaserLoadLimits::default()).unwrap();

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

        assert!(
            scan_slide_laser_traces(0, xml.as_bytes(), &mut LaserLoadLimits::default()).is_err()
        );

        for coordinate in ["27273042316901", "1e2mm", "+1mm", "1px"] {
            let xml = format!(
                r#"<p:sld xmlns:p="{PML}" xmlns:p14="{P14_NAMESPACE}"><p:extLst><p:ext uri="{LASER_TRACE_EXTENSION_URI}"><p14:laserTraceLst><p14:tracePtLst><p14:tracePt t="0" x="{coordinate}" y="0"/></p14:tracePtLst></p14:laserTraceLst></p:ext></p:extLst></p:sld>"#
            );
            assert!(
                scan_slide_laser_traces(0, xml.as_bytes(), &mut LaserLoadLimits::default())
                    .is_err(),
                "accepted {coordinate:?}"
            );
        }
    }

    fn slide_package(tail: &str) -> (litchi_opc::OpcPackage, litchi_opc::PackURI) {
        let mut package = litchi_opc::OpcPackage::new();
        let name = litchi_opc::PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>{tail}</p:sld>"#
        );
        package.add_part(Box::new(litchi_opc::BlobPart::new(
            name.clone(),
            ct::PML_SLIDE.into(),
            xml.into_bytes(),
        )));
        (package, name)
    }

    fn sample_points() -> Vec<PptxLaserTracePoint> {
        vec![
            PptxLaserTracePoint::new(
                Offset::ZERO,
                Coordinate::emu(914_400).unwrap(),
                Coordinate::emu(457_200).unwrap(),
            ),
            PptxLaserTracePoint::new(
                Offset::ms(2500),
                Coordinate::parse("1.25cm").unwrap(),
                Coordinate::from(34),
            ),
        ]
    }

    #[test]
    fn stores_laser_trace_and_discovers_it_round_trip() {
        let (mut package, slide_name) = slide_package("");
        let points = sample_points();
        store_slide_laser_trace(&mut package, &slide_name, &points).unwrap();

        let slide = package.get_part(&slide_name).unwrap();
        let traces = load_slide_laser_traces(0, slide, &mut LaserLoadLimits::default()).unwrap();
        assert_eq!(traces.len(), 1);
        assert_eq!(traces[0].points(), points.as_slice());

        // A second trace on the same slide is rejected (no replacement).
        assert!(store_slide_laser_trace(&mut package, &slide_name, &points).is_err());
    }

    #[test]
    fn stores_laser_trace_into_existing_and_empty_extension_lists() {
        // Existing non-empty extLst.
        let (mut package, slide_name) = slide_package(
            r#"<p:extLst><p:ext uri="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"/></p:extLst>"#,
        );
        store_slide_laser_trace(&mut package, &slide_name, &sample_points()).unwrap();
        let slide = package.get_part(&slide_name).unwrap();
        let xml = String::from_utf8(slide.blob().to_vec()).unwrap();
        assert!(xml.contains("{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"));
        assert_eq!(
            load_slide_laser_traces(0, slide, &mut LaserLoadLimits::default())
                .unwrap()
                .len(),
            1
        );

        // Empty extLst element.
        let (mut package, slide_name) = slide_package("<p:extLst/>");
        store_slide_laser_trace(&mut package, &slide_name, &sample_points()).unwrap();
        let slide = package.get_part(&slide_name).unwrap();
        let traces = load_slide_laser_traces(0, slide, &mut LaserLoadLimits::default()).unwrap();
        assert_eq!(traces.len(), 1);
        assert_eq!(traces[0].point_count(), 2);
    }

    #[test]
    fn stores_laser_trace_in_strict_dialect() {
        let mut package = litchi_opc::OpcPackage::new();
        let name = litchi_opc::PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let xml = r#"<p:sld xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sld>"#;
        package.add_part(Box::new(litchi_opc::BlobPart::new(
            name.clone(),
            ct::PML_SLIDE.into(),
            xml.as_bytes().to_vec(),
        )));
        store_slide_laser_trace(&mut package, &name, &sample_points()).unwrap();
        let slide = package.get_part(&name).unwrap();
        let xml = String::from_utf8(slide.blob().to_vec()).unwrap();
        assert!(xml.contains("xmlns:p=\"http://purl.oclc.org/ooxml/presentationml/main\""));
        assert_eq!(
            load_slide_laser_traces(0, slide, &mut LaserLoadLimits::default())
                .unwrap()
                .len(),
            1
        );
    }

    #[test]
    fn rejects_invalid_laser_storage_inputs() {
        let (mut package, slide_name) = slide_package("");
        // No points.
        assert!(store_slide_laser_trace(&mut package, &slide_name, &[]).is_err());
        // Bad time offsets cannot enter the typed point constructor.
        assert!(Offset::parse("").is_err());
        assert!(Offset::parse("a<b").is_err());
        // Non-slide part.
        let wrong = litchi_opc::PackURI::new("/ppt/presentation.xml").unwrap();
        package.add_part(Box::new(litchi_opc::BlobPart::new(
            wrong.clone(),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                .into(),
            b"<p:presentation/>".to_vec(),
        )));
        assert!(store_slide_laser_trace(&mut package, &wrong, &sample_points()).is_err());
        // Rejection leaves the slide without an extension list.
        let slide = package.get_part(&slide_name).unwrap();
        assert!(
            load_slide_laser_traces(0, slide, &mut LaserLoadLimits::default())
                .unwrap()
                .is_empty()
        );
    }
}
