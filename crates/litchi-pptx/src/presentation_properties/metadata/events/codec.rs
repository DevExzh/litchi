//! Bounded, inert `PowerPoint` slide-show event discovery.
//!
//! Slide-show events are retained as persisted document history only. This
//! module never replays, renders, seeks, pauses, resumes, stops, or otherwise
//! executes a recorded action.

use std::ops::Range;

use super::model::{Draft, Event, Kind, Trigger};
use super::validation::{
    MAX_EVENTS, MAX_SLIDE_XML_BYTES, MAX_TOTAL_SLIDE_XML_BYTES, MAX_XML_DEPTH, MAX_XML_NODES,
    validate_events, validate_extension_uri,
};
use crate::presentation_properties::metadata::is_presentationml_name;
use crate::time::{Offset, ParseError as TimeParseError};
use crate::{Error, Result};
use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};
use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_opc::Part;
use litchi_opc::constants::content_type as ct;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event as XmlEvent};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

/// The `PowerPoint` extension URI that contains persisted slide-show events.
pub const EXTENSION_URI: &str = "{E180D4A7-C9FB-4DFB-919C-405C955672EB}";

const P14_NAMESPACE: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const P14_NAMESPACE_BYTES: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2010/main";
#[derive(Default)]
pub(crate) struct LoadLimits {
    total_slide_xml_bytes: usize,
    event_count: usize,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum ElementKind {
    Other,
    Root,
    ShowEventExtension,
    ShowEventList,
    TriggerEvent,
    PlayEvent,
    StopEvent,
    PauseEvent,
    ResumeEvent,
    SeekEvent,
    NullEvent,
}

impl ElementKind {
    fn is_known(self) -> bool {
        !matches!(self, Self::Other | Self::Root)
    }
}

/// Load bounded, inert slide-show events from one `PresentationML` slide.
pub(crate) fn load(
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut LoadLimits,
) -> Result<Vec<Event>> {
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid(
            "slide-show event discovery requires a PresentationML slide part",
        ));
    }
    limits.add_slide_xml(slide.blob().len())?;
    scan_slide_show_events(slide_index, slide.blob(), limits)
}

impl LoadLimits {
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

    fn add_event(&mut self) -> Result<()> {
        self.event_count = self
            .event_count
            .checked_add(1)
            .ok_or_else(|| limit("slide-show event count"))?;
        if self.event_count > MAX_EVENTS {
            return Err(limit("slide-show event count"));
        }
        Ok(())
    }
}

fn scan_slide_show_events(
    slide_index: usize,
    xml_bytes: &[u8],
    limits: &mut LoadLimits,
) -> Result<Vec<Event>> {
    if xml_bytes.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("slide XML bytes"));
    }

    let mut capabilities = Capabilities::ooxml_baseline();
    capabilities.understand_namespace(P14_NAMESPACE);
    let mce_limits = Limits {
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
    let mut events = Vec::new();
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
            XmlEvent::Start(element) => {
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
                    slide_index,
                    &mut events,
                    limits,
                )?;
                if kind == ElementKind::Root {
                    saw_root = true;
                }
                stack.push(kind);
            },
            XmlEvent::Empty(element) => {
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
                    slide_index,
                    &mut events,
                    limits,
                )?;
                if kind == ElementKind::Root {
                    saw_root = true;
                    closed_root = true;
                }
            },
            XmlEvent::End(element) => {
                let kind = stack
                    .pop()
                    .ok_or_else(|| invalid("invalid slide XML nesting"))?;
                finish_element(kind, &namespace, element.name())?;
                if kind == ElementKind::Root {
                    closed_root = true;
                }
            },
            XmlEvent::Text(text)
                if stack.last().copied().is_some_and(ElementKind::is_known)
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid("slide-show event markup cannot contain text"));
            },
            XmlEvent::CData(text)
                if stack.last().copied().is_some_and(ElementKind::is_known)
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid("slide-show event markup cannot contain text"));
            },
            XmlEvent::GeneralRef(_) if stack.last().copied().is_some_and(ElementKind::is_known) => {
                return Err(invalid(
                    "slide-show event markup cannot contain entity references",
                ));
            },
            XmlEvent::DocType(_) => return Err(invalid("slide XML must not contain a DTD")),
            XmlEvent::PI(_) => {
                return Err(invalid(
                    "slide XML must not contain a processing instruction",
                ));
            },
            XmlEvent::Eof => {
                if !stack.is_empty() || !saw_root || !closed_root {
                    return Err(invalid("unterminated or missing PresentationML slide root"));
                }
                break;
            },
            _ => {},
        }
    }

    Ok(events)
}

#[allow(
    clippy::too_many_arguments,
    reason = "element classifier threads one slot per event-element field"
)]
fn classify_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    parent: ElementKind,
    depth: usize,
    root_seen: bool,
    empty: bool,
    slide_index: usize,
    events: &mut Vec<Event>,
    limits: &mut LoadLimits,
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
        && is_show_event_extension(element, decoder)?
    {
        return if parent.is_known() {
            Err(invalid("slide-show extension has invalid nesting"))
        } else {
            Ok(ElementKind::ShowEventExtension)
        };
    }

    if is_p14_name(namespace, element.name(), b"showEvtLst") {
        return match parent {
            ElementKind::ShowEventExtension => Ok(ElementKind::ShowEventList),
            ElementKind::Other => Ok(ElementKind::Other),
            ElementKind::Root
            | ElementKind::ShowEventList
            | ElementKind::TriggerEvent
            | ElementKind::PlayEvent
            | ElementKind::StopEvent
            | ElementKind::PauseEvent
            | ElementKind::ResumeEvent
            | ElementKind::SeekEvent
            | ElementKind::NullEvent => Err(invalid(
                "showEvtLst must be the direct child of its PowerPoint extension",
            )),
        };
    }

    if let Some(kind) = event_element_kind(namespace, element.name()) {
        if parent != ElementKind::ShowEventList {
            return if parent.is_known() {
                Err(invalid(
                    "slide-show event is outside a PowerPoint showEvtLst element",
                ))
            } else {
                Ok(ElementKind::Other)
            };
        }
        let event = parse_show_event(slide_index, events.len(), element, decoder)?;
        limits.add_event()?;
        events.push(event);
        return Ok(if empty { ElementKind::Other } else { kind });
    }

    if parent.is_known() {
        return Err(invalid(
            "slide-show event extension contains an unsupported child element",
        ));
    }
    Ok(ElementKind::Other)
}

fn finish_element(kind: ElementKind, namespace: &ResolveResult<'_>, name: QName<'_>) -> Result<()> {
    match kind {
        ElementKind::Root if !is_presentationml_name(namespace, name, b"sld") => Err(invalid(
            "slide XML must close with a PresentationML sld element",
        )),
        ElementKind::ShowEventExtension if !is_presentationml_name(namespace, name, b"ext") => {
            Err(invalid("invalid slide-show extension nesting"))
        },
        ElementKind::ShowEventList if !is_p14_name(namespace, name, b"showEvtLst") => {
            Err(invalid("invalid slide-show event-list nesting"))
        },
        ElementKind::TriggerEvent if !is_p14_name(namespace, name, b"triggerEvt") => {
            Err(invalid("invalid trigger-event nesting"))
        },
        ElementKind::PlayEvent if !is_p14_name(namespace, name, b"playEvt") => {
            Err(invalid("invalid play-event nesting"))
        },
        ElementKind::StopEvent if !is_p14_name(namespace, name, b"stopEvt") => {
            Err(invalid("invalid stop-event nesting"))
        },
        ElementKind::PauseEvent if !is_p14_name(namespace, name, b"pauseEvt") => {
            Err(invalid("invalid pause-event nesting"))
        },
        ElementKind::ResumeEvent if !is_p14_name(namespace, name, b"resumeEvt") => {
            Err(invalid("invalid resume-event nesting"))
        },
        ElementKind::SeekEvent if !is_p14_name(namespace, name, b"seekEvt") => {
            Err(invalid("invalid seek-event nesting"))
        },
        ElementKind::NullEvent if !is_p14_name(namespace, name, b"nullEvt") => {
            Err(invalid("invalid null-event nesting"))
        },
        _ => Ok(()),
    }
}

fn event_element_kind(namespace: &ResolveResult<'_>, name: QName<'_>) -> Option<ElementKind> {
    [
        (b"triggerEvt".as_slice(), ElementKind::TriggerEvent),
        (b"playEvt".as_slice(), ElementKind::PlayEvent),
        (b"stopEvt".as_slice(), ElementKind::StopEvent),
        (b"pauseEvt".as_slice(), ElementKind::PauseEvent),
        (b"resumeEvt".as_slice(), ElementKind::ResumeEvent),
        (b"seekEvt".as_slice(), ElementKind::SeekEvent),
        (b"nullEvt".as_slice(), ElementKind::NullEvent),
    ]
    .into_iter()
    .find_map(|(local_name, kind)| is_p14_name(namespace, name, local_name).then_some(kind))
}

fn is_show_event_extension(element: &BytesStart<'_>, decoder: Decoder) -> Result<bool> {
    Ok(unqualified_attribute_value(element, b"uri", decoder)?
        .is_some_and(|uri| validate_extension_uri(&uri, EXTENSION_URI).is_ok()))
}

fn is_p14_name(namespace: &ResolveResult<'_>, name: QName<'_>, local_name: &[u8]) -> bool {
    name.local_name().as_ref() == local_name
        && matches!(
            namespace,
            ResolveResult::Bound(Namespace(value)) if *value == P14_NAMESPACE_BYTES
        )
}

fn parse_show_event(
    slide_index: usize,
    event_index: usize,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<Event> {
    let time = parse_time_offset(required_attribute(element, b"time", decoder)?)?;
    let object_id = parse_object_id(required_attribute(element, b"objId", decoder)?)?;
    let kind = match element.local_name().as_ref() {
        b"triggerEvt" => Kind::Trigger(parse_trigger(required_attribute(
            element, b"type", decoder,
        )?)?),
        b"playEvt" => Kind::Play,
        b"stopEvt" => Kind::Stop,
        b"pauseEvt" => Kind::Pause,
        b"resumeEvt" => Kind::Resume,
        b"seekEvt" => {
            let at = parse_time_offset(required_attribute(element, b"seek", decoder)?)?;
            Kind::Seek { at }
        },
        b"nullEvt" => Kind::Null,
        _ => return Err(invalid("unsupported slide-show event element")),
    };
    Ok(Event {
        slide_index,
        event_index,
        kind,
        time,
        object_id,
    })
}

fn required_attribute(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<String> {
    unqualified_attribute_value(element, name, decoder)?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            invalid(format!(
                "slide-show event is missing required '{}' attribute",
                String::from_utf8_lossy(name)
            ))
        })
}

fn parse_object_id(value: String) -> Result<u32> {
    value
        .parse()
        .map_err(|_err| invalid("slide-show event object ID must be an unsigned 32-bit integer"))
}

fn parse_trigger(value: String) -> Result<Trigger> {
    match value.as_str() {
        "none" => Ok(Trigger::None),
        "onBegin" => Ok(Trigger::OnBegin),
        "onEnd" => Ok(Trigger::OnEnd),
        "begin" => Ok(Trigger::Begin),
        "end" => Ok(Trigger::End),
        "onClick" => Ok(Trigger::OnClick),
        "onDblClick" => Ok(Trigger::OnDoubleClick),
        "onMouseOver" => Ok(Trigger::OnMouseOver),
        "onMouseOut" => Ok(Trigger::OnMouseOut),
        "onNext" => Ok(Trigger::OnNext),
        "onPrev" => Ok(Trigger::OnPrevious),
        "onStopAudio" => Ok(Trigger::OnStopAudio),
        _ => Err(invalid(format!(
            "unsupported slide-show trigger type '{value}'"
        ))),
    }
}

fn parse_time_offset(value: String) -> Result<Offset> {
    Offset::try_from(value).map_err(time_error)
}

fn time_error(error: TimeParseError) -> Error {
    invalid(format!(
        "invalid slide-show event universal time offset: {error}"
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

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(what: &str) -> Error {
    invalid(format!("{what} exceeds the supported safety limit"))
}

fn trigger_token(trigger: Trigger) -> &'static str {
    match trigger {
        Trigger::None => "none",
        Trigger::OnBegin => "onBegin",
        Trigger::OnEnd => "onEnd",
        Trigger::Begin => "begin",
        Trigger::End => "end",
        Trigger::OnClick => "onClick",
        Trigger::OnDoubleClick => "onDblClick",
        Trigger::OnMouseOver => "onMouseOver",
        Trigger::OnMouseOut => "onMouseOut",
        Trigger::OnNext => "onNext",
        Trigger::OnPrevious => "onPrev",
        Trigger::OnStopAudio => "onStopAudio",
    }
}

impl Draft {
    pub(crate) fn element_name(&self) -> &'static str {
        match &self.kind {
            Kind::Trigger(_) => "triggerEvt",
            Kind::Play => "playEvt",
            Kind::Stop => "stopEvt",
            Kind::Pause => "pauseEvt",
            Kind::Resume => "resumeEvt",
            Kind::Seek { .. } => "seekEvt",
            Kind::Null => "nullEvt",
        }
    }
}

/// Store slide-show event records onto a slide as a `PowerPoint` 2010
/// `p14:showEvtLst` extension.
///
/// The typed events are serialized canonically in caller order; the
/// slide gains the `p:ext` extension block (patched into an existing
/// extension list, expanding an empty one, or creating one) while
/// preserving its namespace dialect. Slides that already carry a show-event
/// extension are rejected — replacement is not supported in this pass.
/// Events are never replayed, rendered, or executed.
pub(super) fn store(
    package: &mut litchi_opc::OpcPackage,
    slide_name: &litchi_opc::PackURI,
    events: &[Draft],
) -> Result<()> {
    validate_events(events)?;
    let slide = package.get_part(slide_name)?;
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid(
            "slide-show event storage requires a PresentationML slide part",
        ));
    }
    if !load(0, slide, &mut LoadLimits::default())?.is_empty() {
        return Err(invalid(
            "slide already contains a slide-show event extension; replacement is not supported",
        ));
    }

    let mut fragment = String::with_capacity(events.len() * 72 + 256);
    fragment.push_str("<p:ext xmlns:p=\"");
    fragment.push_str(super::super::slide_patch::slide_dialect(slide.blob())?);
    fragment.push_str("\" xmlns:p14=\"");
    fragment.push_str(P14_NAMESPACE);
    fragment.push_str("\" uri=\"");
    fragment.push_str(EXTENSION_URI);
    fragment.push_str("\"><p14:showEvtLst>");
    for event in events {
        fragment.push_str("<p14:");
        fragment.push_str(event.element_name());
        if let Kind::Trigger(trigger) = &event.kind {
            fragment.push_str(" type=\"");
            fragment.push_str(trigger_token(*trigger));
            fragment.push('"');
        }
        fragment.push_str(" time=\"");
        fragment.push_str(event.time.as_str());
        fragment.push_str("\" objId=\"");
        fragment.push_str(&event.object_id.to_string());
        fragment.push('"');
        if let Kind::Seek { at: seek_time } = &event.kind {
            fragment.push_str(" seek=\"");
            fragment.push_str(seek_time.as_str());
            fragment.push('"');
        }
        fragment.push_str("/>");
    }
    fragment.push_str("</p14:showEvtLst></p:ext>");

    let updated = super::super::slide_patch::insert_extension_fragment(slide.blob(), &fragment)?;
    // Self-check: the patched slide must read back through discovery.
    let probe =
        litchi_opc::BlobPart::new(slide_name.clone(), ct::PML_SLIDE.into(), updated.clone());
    let discovered = load(0, &probe, &mut LoadLimits::default())?;
    if discovered.len() != events.len() {
        return Err(invalid(
            "slide-show event storage failed read-back validation",
        ));
    }
    package.get_part_mut(slide_name)?.set_blob(updated);
    Ok(())
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct AttributeSpan {
    name: Vec<u8>,
    name_start: usize,
    value_start: usize,
    value_end: usize,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct EventSpan {
    range: Range<usize>,
    open_end: usize,
    qualified_name: Range<usize>,
    local_name: Range<usize>,
    close_local_name: Option<Range<usize>>,
    attributes: Vec<AttributeSpan>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct Located {
    pub(crate) events: Vec<Event>,
    pub(crate) spans: Vec<EventSpan>,
    pub(crate) extension: Range<usize>,
    pub(crate) list_content: Range<usize>,
}

#[derive(Clone, Debug)]
enum Frame {
    Root,
    Other,
    Extension { start: usize },
    List { open_end: usize },
    Event(EventSpan),
}

#[derive(Clone, Debug)]
struct Edit {
    range: Range<usize>,
    replacement: Vec<u8>,
}

/// Capture the event records and exact source ranges needed for a detached
/// transaction. The source XML is never executed, resolved, or dereferenced.
pub(crate) fn locate(
    slide_index: usize,
    source: &[u8],
    limits: &mut LoadLimits,
) -> Result<Option<Located>> {
    limits.add_slide_xml(source.len())?;
    let events = scan_slide_show_events(slide_index, source, limits)?;
    if events.is_empty() {
        return Ok(None);
    }
    let mut located = locate_raw(source)?;
    if located.spans.len() != events.len() {
        return Err(invalid(
            "slide-show event source ranges differ from the validated event list",
        ));
    }
    located.events = events;
    Ok(Some(located))
}

fn locate_raw(source: &[u8]) -> Result<Located> {
    if source.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("slide XML bytes"));
    }
    let mut reader = NsReader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut stack = Vec::new();
    let mut spans = Vec::new();
    let mut extension = None;
    let mut list_content = None;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut extension_count = 0usize;
    let mut list_count = 0usize;
    let mut nodes = 0usize;

    loop {
        let before = usize::try_from(reader.buffer_position())
            .map_err(|_err| invalid("slide XML offset overflow"))?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let after = usize::try_from(reader.buffer_position())
            .map_err(|_err| invalid("slide XML offset overflow"))?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            XmlEvent::Start(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit("slide XML node count"))?;
                if nodes > MAX_XML_NODES || stack.len() + 1 > MAX_XML_DEPTH {
                    return Err(limit("slide XML resources"));
                }
                let parent = stack.last();
                let kind = if stack.is_empty() {
                    if root_seen || !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid(
                            "slide XML must have one PresentationML sld root element",
                        ));
                    }
                    root_seen = true;
                    Frame::Root
                } else if is_presentationml_name(&namespace, element.name(), b"ext")
                    && is_show_event_extension(&element, decoder)?
                {
                    extension_count += 1;
                    if extension_count > 1 {
                        return Err(invalid("slide has multiple slide-show event extensions"));
                    }
                    Frame::Extension { start: before }
                } else if is_p14_name(&namespace, element.name(), b"showEvtLst")
                    && parent.is_some_and(|frame| matches!(frame, Frame::Extension { .. }))
                {
                    list_count += 1;
                    if list_count > 1 {
                        return Err(invalid("slide has multiple slide-show event lists"));
                    }
                    Frame::List { open_end: after }
                } else if event_element_kind(&namespace, element.name()).is_some()
                    && parent.is_some_and(|frame| matches!(frame, Frame::List { .. }))
                {
                    Frame::Event(make_event_span(source, before, after)?)
                } else {
                    Frame::Other
                };
                stack.push(kind);
            },
            XmlEvent::Empty(element) => {
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit("slide XML node count"))?;
                if nodes > MAX_XML_NODES || stack.len() + 1 > MAX_XML_DEPTH {
                    return Err(limit("slide XML resources"));
                }
                let parent = stack.last();
                if stack.is_empty() {
                    if root_seen || !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid(
                            "slide XML must have one PresentationML sld root element",
                        ));
                    }
                    root_seen = true;
                    root_closed = true;
                } else if is_presentationml_name(&namespace, element.name(), b"ext")
                    && is_show_event_extension(&element, decoder)?
                {
                    extension_count += 1;
                    if extension_count > 1 {
                        return Err(invalid("slide has multiple slide-show event extensions"));
                    }
                    extension = Some(before..after);
                } else if is_p14_name(&namespace, element.name(), b"showEvtLst")
                    && parent.is_some_and(|frame| matches!(frame, Frame::Extension { .. }))
                {
                    list_count += 1;
                    if list_count > 1 {
                        return Err(invalid("slide has multiple slide-show event lists"));
                    }
                    list_content = Some(after..before);
                } else if event_element_kind(&namespace, element.name()).is_some()
                    && parent.is_some_and(|frame| matches!(frame, Frame::List { .. }))
                {
                    spans.push(make_event_span(source, before, after)?);
                }
            },
            XmlEvent::End(element) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("invalid slide XML nesting"))?;
                match frame {
                    Frame::Root => {
                        if !is_presentationml_name(&namespace, element.name(), b"sld") {
                            return Err(invalid("slide XML root closes with the wrong element"));
                        }
                        root_closed = true;
                    },
                    Frame::Extension { start } => {
                        extension = Some(start..after);
                    },
                    Frame::List { open_end } => {
                        list_content = Some(open_end..before);
                    },
                    Frame::Event(mut span) => {
                        span.close_local_name =
                            Some(element_name_ranges(source, before, after, true)?.1);
                        span.range.end = after;
                        spans.push(span);
                    },
                    Frame::Other => {},
                }
            },
            XmlEvent::DocType(_) | XmlEvent::PI(_) => {
                return Err(invalid(
                    "slide XML must not contain a DTD or processing instruction",
                ));
            },
            XmlEvent::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() || !root_seen || !root_closed {
        return Err(invalid("unterminated or missing PresentationML slide root"));
    }
    spans.sort_by_key(|span| span.range.start);
    let Some(extension) = extension else {
        return Err(invalid(
            "validated slide-show events have no extension source",
        ));
    };
    let Some(list_content) = list_content else {
        return Err(invalid(
            "validated slide-show events have no event-list source",
        ));
    };
    Ok(Located {
        events: Vec::new(),
        spans,
        extension,
        list_content,
    })
}

fn make_event_span(source: &[u8], start: usize, open_end: usize) -> Result<EventSpan> {
    let (qualified_name, local_name) = element_name_ranges(source, start, open_end, false)?;
    Ok(EventSpan {
        range: start..open_end,
        open_end,
        qualified_name,
        local_name,
        close_local_name: None,
        attributes: scan_attributes(source, start, open_end)?,
    })
}

fn element_name_ranges(
    source: &[u8],
    start: usize,
    open_end: usize,
    closing: bool,
) -> Result<(Range<usize>, Range<usize>)> {
    let mut cursor = start
        .checked_add(if closing { 2 } else { 1 })
        .ok_or_else(|| invalid("slide XML offset overflow"))?;
    let name_start = cursor;
    while cursor < open_end
        && !is_xml_space(source[cursor])
        && !matches!(source[cursor], b'/' | b'>')
    {
        cursor += 1;
    }
    if cursor == name_start {
        return Err(invalid("slide XML element has no name"));
    }
    let local_start = source[name_start..cursor]
        .iter()
        .rposition(|byte| *byte == b':')
        .map_or(name_start, |offset| name_start + offset + 1);
    Ok((name_start..cursor, local_start..cursor))
}

fn scan_attributes(source: &[u8], start: usize, open_end: usize) -> Result<Vec<AttributeSpan>> {
    let end = open_end
        .checked_sub(1)
        .ok_or_else(|| invalid("slide XML opening tag is truncated"))?;
    let mut cursor = start + 1;
    while cursor < end && !is_xml_space(source[cursor]) && source[cursor] != b'/' {
        cursor += 1;
    }
    let mut attributes = Vec::new();
    while cursor < end {
        while cursor < end && is_xml_space(source[cursor]) {
            cursor += 1;
        }
        if cursor >= end || source[cursor] == b'/' {
            break;
        }
        let name_start = cursor;
        while cursor < end
            && !is_xml_space(source[cursor])
            && !matches!(source[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        if name_start == cursor {
            return Err(invalid("slide XML attribute has no name"));
        }
        let name = source[name_start..cursor].to_vec();
        while cursor < end && is_xml_space(source[cursor]) {
            cursor += 1;
        }
        if cursor >= end || source[cursor] != b'=' {
            return Err(invalid("slide XML attribute has no value"));
        }
        cursor += 1;
        while cursor < end && is_xml_space(source[cursor]) {
            cursor += 1;
        }
        if cursor >= end || !matches!(source[cursor], b'\'' | b'"') {
            return Err(invalid("slide XML attribute value is not quoted"));
        }
        let quote = source[cursor];
        cursor += 1;
        let value_start = cursor;
        while cursor < end && source[cursor] != quote {
            cursor += 1;
        }
        if cursor >= end {
            return Err(invalid("slide XML attribute value is unterminated"));
        }
        attributes.push(AttributeSpan {
            name,
            name_start,
            value_start,
            value_end: cursor,
        });
        cursor += 1;
    }
    Ok(attributes)
}

/// Rewrite modeled event attributes while retaining unknown attributes and
/// all bytes outside the event list.
pub(crate) fn rewrite(source: &[u8], located: &Located, working: &[Draft]) -> Result<Vec<u8>> {
    validate_events(working)?;
    if located.spans.len() != located.events.len() {
        return Err(invalid("event source span count is inconsistent"));
    }
    if working.len() != located.spans.len() {
        let prefix = event_prefix(source, &located.spans[0], b"p14:")?;
        let mut body = Vec::with_capacity(working.len() * 72);
        for event in working {
            body.extend_from_slice(&serialize_event(event, &prefix));
        }
        return apply_edits(
            source,
            vec![Edit {
                range: located.list_content.clone(),
                replacement: body,
            }],
        );
    }

    let mut edits = Vec::new();
    for ((source_event, span), desired) in located.events.iter().zip(&located.spans).zip(working) {
        let original = Draft::from_event(source_event);
        if original == *desired {
            continue;
        }
        edits.extend(rewrite_event(source, span, desired)?);
    }
    apply_edits(source, edits)
}

pub(crate) fn remove_extension(source: &[u8], located: &Located) -> Result<Vec<u8>> {
    apply_edits(
        source,
        vec![Edit {
            range: located.extension.clone(),
            replacement: Vec::new(),
        }],
    )
}

fn rewrite_event(source: &[u8], span: &EventSpan, desired: &Draft) -> Result<Vec<Edit>> {
    let mut edits = Vec::new();
    let new_local = desired.element_name().as_bytes();
    edits.push(Edit {
        range: span.local_name.clone(),
        replacement: new_local.to_vec(),
    });
    if let Some(close_local_name) = &span.close_local_name {
        edits.push(Edit {
            range: close_local_name.clone(),
            replacement: new_local.to_vec(),
        });
    }
    let time = desired.time().as_str().to_owned();
    edits.push(attribute_edit(source, span, b"time", Some(&time))?);
    edits.push(attribute_edit(
        source,
        span,
        b"objId",
        Some(&desired.object_id().to_string()),
    )?);
    let trigger = match desired.kind() {
        Kind::Trigger(value) => Some(trigger_token(*value).to_owned()),
        _ => None,
    };
    edits.push(attribute_edit(source, span, b"type", trigger.as_deref())?);
    let seek = match desired.kind() {
        Kind::Seek { at } => Some(at.as_str().to_owned()),
        _ => None,
    };
    edits.push(attribute_edit(source, span, b"seek", seek.as_deref())?);
    Ok(edits)
}

fn attribute_edit(
    source: &[u8],
    span: &EventSpan,
    name: &[u8],
    value: Option<&str>,
) -> Result<Edit> {
    let matches: Vec<_> = span
        .attributes
        .iter()
        .filter(|attribute| attribute.name.as_slice() == name)
        .collect();
    if matches.len() > 1 {
        return Err(invalid(format!(
            "slide-show event contains duplicate '{}' attributes",
            String::from_utf8_lossy(name)
        )));
    }
    match (matches.first().copied(), value) {
        (Some(attribute), Some(value)) => Ok(Edit {
            range: attribute.value_start..attribute.value_end,
            replacement: crate::presentation_properties::metadata::escape_xml(value).into_bytes(),
        }),
        (Some(attribute), None) => {
            let mut start = attribute.name_start;
            while start > span.qualified_name.end && is_xml_space(source[start - 1]) {
                start -= 1;
            }
            Ok(Edit {
                range: start..attribute.value_end + 1,
                replacement: Vec::new(),
            })
        },
        (None, Some(value)) => {
            let mut at = span.open_end - 1;
            if at > span.qualified_name.end && source[at - 1] == b'/' {
                at -= 1;
            }
            Ok(Edit {
                range: at..at,
                replacement: format!(
                    " {}=\"{}\"",
                    String::from_utf8_lossy(name),
                    crate::presentation_properties::metadata::escape_xml(value)
                )
                .into_bytes(),
            })
        },
        (None, None) => Ok(Edit {
            range: 0..0,
            replacement: Vec::new(),
        }),
    }
}

fn event_prefix(source: &[u8], span: &EventSpan, fallback: &[u8]) -> Result<Vec<u8>> {
    let qualified = &source[span.qualified_name.clone()];
    let local = &source[span.local_name.clone()];
    let prefix_len = qualified
        .len()
        .checked_sub(local.len())
        .ok_or_else(|| invalid("slide-show event name range is inconsistent"))?;
    if prefix_len <= qualified.len() {
        Ok(qualified[..prefix_len].to_vec())
    } else {
        Ok(fallback.to_vec())
    }
}

fn serialize_event(event: &Draft, prefix: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(96);
    out.extend_from_slice(b"<");
    out.extend_from_slice(prefix);
    out.extend_from_slice(event.element_name().as_bytes());
    if let Kind::Trigger(trigger) = event.kind() {
        out.extend_from_slice(b" type=\"");
        out.extend_from_slice(trigger_token(*trigger).as_bytes());
        out.extend_from_slice(b"\"");
    }
    out.extend_from_slice(b" time=\"");
    out.extend_from_slice(event.time().as_str().as_bytes());
    out.extend_from_slice(b"\" objId=\"");
    out.extend_from_slice(event.object_id().to_string().as_bytes());
    out.extend_from_slice(b"\"");
    if let Kind::Seek { at } = event.kind() {
        out.extend_from_slice(b" seek=\"");
        out.extend_from_slice(at.as_str().as_bytes());
        out.extend_from_slice(b"\"");
    }
    out.extend_from_slice(b"/>");
    out
}

fn apply_edits(source: &[u8], mut edits: Vec<Edit>) -> Result<Vec<u8>> {
    edits.retain(|edit| edit.range.start != edit.range.end || !edit.replacement.is_empty());
    edits.sort_by(|left, right| {
        right
            .range
            .start
            .cmp(&left.range.start)
            .then_with(|| right.range.end.cmp(&left.range.end))
    });
    for pair in edits.windows(2) {
        if pair[0].range.start < pair[1].range.end {
            return Err(invalid("slide-show source edits overlap"));
        }
    }
    let mut output = source.to_vec();
    for edit in edits {
        if edit.range.end > output.len() {
            return Err(invalid("slide-show source edit escapes the slide"));
        }
        output.splice(edit.range, edit.replacement);
    }
    if output.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("updated slide XML bytes"));
    }
    Ok(output)
}

fn is_xml_space(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\n' | b'\r')
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
    fn scans_show_events_through_markup_compatibility() {
        let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:mc="{MCE}" xmlns:p14="{P14_NAMESPACE}" mc:Ignorable="p14"><p:extLst><p:ext uri="{EXTENSION_URI}"><p14:showEvtLst><p14:triggerEvt type="onClick" time="1.5s" objId="6"/><p14:seekEvt time="2000ms" objId="4" seek="10.379s"/><p14:nullEvt time="3" objId="5"/></p14:showEvtLst></p:ext></p:extLst></p:sld>"#
        );
        let events = scan_slide_show_events(4, xml.as_bytes(), &mut LoadLimits::default()).unwrap();

        assert_eq!(events.len(), 3);
        assert_eq!(events[0].slide_index(), 4);
        assert_eq!(events[0].event_index(), 0);
        assert_eq!(events[0].kind(), &Kind::Trigger(Trigger::OnClick));
        assert_eq!(events[0].time(), &Offset::parse("1.5s").unwrap());
        assert_eq!(events[0].object_id(), 6);
        assert_eq!(events[0].seek_time(), None);
        assert!(matches!(
            events[1].kind(),
            Kind::Seek { at }
                if at == &Offset::parse("10.379s").unwrap()
        ));
        assert_eq!(
            events[1].seek_time(),
            Some(&Offset::parse("10.379s").unwrap())
        );
        assert_eq!(events[2].kind(), &Kind::Null);
    }

    #[test]
    fn rejects_malformed_show_events() {
        let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:p14="{P14_NAMESPACE}"><p:extLst><p:ext uri="{EXTENSION_URI}"><p14:showEvtLst><p14:seekEvt time="1..2s" objId="4" seek="0"/></p14:showEvtLst></p:ext></p:extLst></p:sld>"#
        );

        assert!(scan_slide_show_events(0, xml.as_bytes(), &mut LoadLimits::default()).is_err());
    }

    #[test]
    fn retains_office_none_trigger_and_rejects_non_domain_bookmark_trigger() {
        let accepted = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:p14="{P14_NAMESPACE}"><p:extLst><p:ext uri="{EXTENSION_URI}"><p14:showEvtLst><p14:triggerEvt type="none" time="0" objId="1"/></p14:showEvtLst></p:ext></p:extLst></p:sld>"#
        );
        let events =
            scan_slide_show_events(0, accepted.as_bytes(), &mut LoadLimits::default()).unwrap();
        assert_eq!(events[0].kind(), &Kind::Trigger(Trigger::None));

        let rejected = accepted.replace("type=\"none\"", "type=\"onMediaBookmark\"");
        assert!(
            scan_slide_show_events(0, rejected.as_bytes(), &mut LoadLimits::default(),).is_err()
        );
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

    fn sample_events() -> Vec<Draft> {
        vec![
            Draft::trigger(Trigger::OnClick, Offset::parse("1.5s").unwrap(), 6),
            Draft::play(Offset::secs(2), 4),
            Draft::seek(Offset::ms(2500), 4, Offset::parse("10.379s").unwrap()),
            Draft::null(Offset::ms(3), 5),
        ]
    }

    #[test]
    fn stores_show_events_and_discovers_them_round_trip() {
        let (mut package, slide_name) = slide_package("");
        let events = sample_events();
        store(&mut package, &slide_name, &events).unwrap();

        let slide = package.get_part(&slide_name).unwrap();
        let discovered = load(0, slide, &mut LoadLimits::default()).unwrap();
        assert_eq!(discovered.len(), events.len());
        assert_eq!(discovered[0].kind(), &Kind::Trigger(Trigger::OnClick));
        assert_eq!(discovered[0].time(), &Offset::ms(1500));
        assert_eq!(discovered[0].object_id(), 6);
        assert_eq!(discovered[1].kind(), &Kind::Play);
        assert!(matches!(
            discovered[2].kind(),
            Kind::Seek { at }
                if at == &Offset::parse("10.379s").unwrap()
        ));
        assert_eq!(
            discovered[2].seek_time(),
            Some(&Offset::parse("10.379s").unwrap())
        );
        assert_eq!(discovered[3].kind(), &Kind::Null);

        // A second storage on the same slide is rejected (no replacement).
        assert!(store(&mut package, &slide_name, &events).is_err());
    }

    #[test]
    fn stores_show_events_into_existing_extension_list() {
        let (mut package, slide_name) = slide_package(
            r#"<p:extLst><p:ext uri="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"/></p:extLst>"#,
        );
        store(&mut package, &slide_name, &sample_events()).unwrap();
        let slide = package.get_part(&slide_name).unwrap();
        let xml = String::from_utf8(slide.blob().to_vec()).unwrap();
        assert!(xml.contains("{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"));
        assert!(xml.contains("p14:showEvtLst"));
        assert_eq!(load(0, slide, &mut LoadLimits::default()).unwrap().len(), 4);
    }

    #[test]
    fn rejects_invalid_show_event_storage() {
        let (mut package, slide_name) = slide_package("");
        // No events.
        assert!(store(&mut package, &slide_name, &[]).is_err());
        // Seek owns its offset in the action variant; non-seek actions cannot
        // carry one in memory or emit one during serialization.
        let seek = Draft::seek(Offset::secs(1), 1, Offset::ms(250));
        assert!(matches!(
            seek.kind(),
            Kind::Seek { at } if at == &Offset::ms(250)
        ));
        assert_eq!(Draft::play(Offset::secs(1), 1).seek_time(), None);
        assert!(Offset::parse("1..2s").is_err());
        assert!(Offset::parse("bad!!").is_err());
        // Non-slide part.
        let wrong = litchi_opc::PackURI::new("/ppt/presentation.xml").unwrap();
        package.add_part(Box::new(litchi_opc::BlobPart::new(
            wrong.clone(),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                .into(),
            b"<p:presentation/>".to_vec(),
        )));
        assert!(store(&mut package, &wrong, &sample_events()).is_err());
        // Rejection leaves the slide without an extension list.
        let slide = package.get_part(&slide_name).unwrap();
        assert!(
            load(0, slide, &mut LoadLimits::default())
                .unwrap()
                .is_empty()
        );
    }
}
