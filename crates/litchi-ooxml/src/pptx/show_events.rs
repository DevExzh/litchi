//! Bounded, inert PowerPoint slide-show event discovery.
//!
//! Slide-show events are retained as persisted document history only. This
//! module never replays, renders, seeks, pauses, resumes, stops, or otherwise
//! executes a recorded action.

use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::is_presentationml_name;
use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_ooxml_common::{MceCapabilities, MceLimits, process_markup_compatibility};
use litchi_opc::Part;
use litchi_opc::constants::content_type as ct;
use litchi_pptx::time::{Offset, ParseError as TimeParseError};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event as XmlEvent};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

/// The PowerPoint extension URI that contains persisted slide-show events.
pub const SHOW_EVENT_EXTENSION_URI: &str = "{E180D4A7-C9FB-4DFB-919C-405C955672EB}";

const P14_NAMESPACE: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const P14_NAMESPACE_BYTES: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2010/main";
const MAX_SLIDE_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_TOTAL_SLIDE_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_SHOW_EVENTS: usize = 65_536;
const MAX_XML_NODES: usize = 250_000;
const MAX_XML_DEPTH: usize = 128;

/// A trigger type recorded by a PowerPoint slide show.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Trigger {
    /// Office extension accepted by PowerPoint 2007 through 2007 SP2; see
    /// `[MS-OE376]` section 2.1.1237.
    None,
    OnBegin,
    OnEnd,
    Begin,
    End,
    OnClick,
    OnDoubleClick,
    OnMouseOver,
    OnMouseOut,
    OnNext,
    OnPrevious,
    OnStopAudio,
}

/// The recorded action represented by a PowerPoint slide-show event.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum EventKind {
    Trigger(Trigger),
    Play,
    Stop,
    Pause,
    Resume,
    /// Seek the targeted media object to an exact stream offset.
    Seek {
        at: Offset,
    },
    /// A reserved unknown event record for future PowerPoint extensions.
    Null,
}

/// A bounded, inert event record persisted for a PowerPoint slide show.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Event {
    slide_index: usize,
    event_index: usize,
    kind: EventKind,
    time: Offset,
    object_id: u32,
}

impl Event {
    /// Return the zero-based index of the slide that owns this event.
    #[inline]
    pub fn slide_index(&self) -> usize {
        self.slide_index
    }

    /// Return the zero-based source-order index of this event on its slide.
    #[inline]
    pub fn event_index(&self) -> usize {
        self.event_index
    }

    /// Return the recorded event kind.
    #[inline]
    pub fn kind(&self) -> &EventKind {
        &self.kind
    }

    /// Return the exact normalized time offset in the slide timeline.
    #[inline]
    pub fn time(&self) -> &Offset {
        &self.time
    }

    /// Return the DrawingML object identifier targeted by this event.
    #[inline]
    pub fn object_id(&self) -> u32 {
        self.object_id
    }

    /// Return the exact normalized media-stream offset for a seek event.
    #[inline]
    pub fn seek_time(&self) -> Option<&Offset> {
        match &self.kind {
            EventKind::Seek { at } => Some(at),
            _ => None,
        }
    }
}

#[derive(Default)]
pub(crate) struct ShowEventLoadLimits {
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

/// Load bounded, inert slide-show events from one PresentationML slide.
pub(crate) fn load_slide_show_events(
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut ShowEventLoadLimits,
) -> Result<Vec<Event>> {
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid(
            "slide-show event discovery requires a PresentationML slide part",
        ));
    }
    limits.add_slide_xml(slide.blob().len())?;
    scan_slide_show_events(slide_index, slide.blob(), limits)
}

impl ShowEventLoadLimits {
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
        if self.event_count > MAX_SHOW_EVENTS {
            return Err(limit("slide-show event count"));
        }
        Ok(())
    }
}

fn scan_slide_show_events(
    slide_index: usize,
    xml_bytes: &[u8],
    limits: &mut ShowEventLoadLimits,
) -> Result<Vec<Event>> {
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
    let mut events = Vec::new();
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

#[allow(clippy::too_many_arguments)]
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
    limits: &mut ShowEventLoadLimits,
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
    Ok(
        unqualified_attribute_value(element, b"uri", decoder)?.as_deref()
            == Some(SHOW_EVENT_EXTENSION_URI),
    )
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
        b"triggerEvt" => EventKind::Trigger(parse_trigger(required_attribute(
            element, b"type", decoder,
        )?)?),
        b"playEvt" => EventKind::Play,
        b"stopEvt" => EventKind::Stop,
        b"pauseEvt" => EventKind::Pause,
        b"resumeEvt" => EventKind::Resume,
        b"seekEvt" => {
            let at = parse_time_offset(required_attribute(element, b"seek", decoder)?)?;
            EventKind::Seek { at }
        },
        b"nullEvt" => EventKind::Null,
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
        .map_err(|_| invalid("slide-show event object ID must be an unsigned 32-bit integer"))
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

fn time_error(error: TimeParseError) -> OoxmlError {
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

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn limit(what: &str) -> OoxmlError {
    invalid(format!("{what} exceeds the supported safety limit"))
}

/// A slide-show event ready for storage onto a slide.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EventDraft {
    kind: EventKind,
    time: Offset,
    object_id: u32,
}

impl EventDraft {
    /// Create an event from its complete, typed action state.
    #[must_use]
    pub fn new(kind: EventKind, time: Offset, object_id: u32) -> Self {
        Self {
            kind,
            time,
            object_id,
        }
    }

    /// Create a trigger event.
    #[must_use]
    pub fn trigger(trigger: Trigger, time: Offset, object_id: u32) -> Self {
        Self::new(EventKind::Trigger(trigger), time, object_id)
    }

    /// Create a play event.
    #[must_use]
    pub fn play(time: Offset, object_id: u32) -> Self {
        Self::new(EventKind::Play, time, object_id)
    }

    /// Create a stop event.
    #[must_use]
    pub fn stop(time: Offset, object_id: u32) -> Self {
        Self::new(EventKind::Stop, time, object_id)
    }

    /// Create a pause event.
    #[must_use]
    pub fn pause(time: Offset, object_id: u32) -> Self {
        Self::new(EventKind::Pause, time, object_id)
    }

    /// Create a resume event.
    #[must_use]
    pub fn resume(time: Offset, object_id: u32) -> Self {
        Self::new(EventKind::Resume, time, object_id)
    }

    /// Create a seek event with a media-stream offset.
    #[must_use]
    pub fn seek(time: Offset, object_id: u32, seek_time: Offset) -> Self {
        Self::new(EventKind::Seek { at: seek_time }, time, object_id)
    }

    /// Create a reserved null event.
    #[must_use]
    pub fn null(time: Offset, object_id: u32) -> Self {
        Self::new(EventKind::Null, time, object_id)
    }

    /// Return the recorded event kind.
    pub fn kind(&self) -> &EventKind {
        &self.kind
    }

    /// Return the exact normalized universal time offset.
    pub fn time(&self) -> &Offset {
        &self.time
    }

    /// Return the DrawingML object identifier targeted by this event.
    pub fn object_id(&self) -> u32 {
        self.object_id
    }

    /// Return the exact normalized media-stream offset for a seek event.
    pub fn seek_time(&self) -> Option<&Offset> {
        match &self.kind {
            EventKind::Seek { at } => Some(at),
            _ => None,
        }
    }

    fn element_name(&self) -> &'static str {
        match &self.kind {
            EventKind::Trigger(_) => "triggerEvt",
            EventKind::Play => "playEvt",
            EventKind::Stop => "stopEvt",
            EventKind::Pause => "pauseEvt",
            EventKind::Resume => "resumeEvt",
            EventKind::Seek { .. } => "seekEvt",
            EventKind::Null => "nullEvt",
        }
    }
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

/// Store slide-show event records onto a slide as a PowerPoint 2010
/// `p14:showEvtLst` extension.
///
/// The typed events are serialized canonically in caller order; the
/// slide gains the `p:ext` extension block (patched into an existing
/// extension list, expanding an empty one, or creating one) while
/// preserving its namespace dialect. Slides that already carry a show-event
/// extension are rejected — replacement is not supported in this pass.
/// Events are never replayed, rendered, or executed.
pub fn store_slide_show_events(
    package: &mut litchi_opc::OpcPackage,
    slide_name: &litchi_opc::PackURI,
    events: &[EventDraft],
) -> Result<()> {
    if events.is_empty() {
        return Err(invalid(
            "slide-show event storage requires at least one event",
        ));
    }
    if events.len() > MAX_SHOW_EVENTS {
        return Err(limit("slide-show event count"));
    }
    let slide = package.get_part(slide_name)?;
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid(
            "slide-show event storage requires a PresentationML slide part",
        ));
    }
    if !load_slide_show_events(0, slide, &mut ShowEventLoadLimits::default())?.is_empty() {
        return Err(invalid(
            "slide already contains a slide-show event extension; replacement is not supported",
        ));
    }

    let mut fragment = String::with_capacity(events.len() * 72 + 256);
    fragment.push_str("<p:ext xmlns:p=\"");
    fragment.push_str(crate::pptx::slide_patch::slide_dialect(slide.blob())?);
    fragment.push_str("\" xmlns:p14=\"");
    fragment.push_str(P14_NAMESPACE);
    fragment.push_str("\" uri=\"");
    fragment.push_str(SHOW_EVENT_EXTENSION_URI);
    fragment.push_str("\"><p14:showEvtLst>");
    for event in events {
        fragment.push_str("<p14:");
        fragment.push_str(event.element_name());
        if let EventKind::Trigger(trigger) = &event.kind {
            fragment.push_str(" type=\"");
            fragment.push_str(trigger_token(*trigger));
            fragment.push('"');
        }
        fragment.push_str(" time=\"");
        fragment.push_str(event.time.as_str());
        fragment.push_str("\" objId=\"");
        fragment.push_str(&event.object_id.to_string());
        fragment.push('"');
        if let EventKind::Seek { at: seek_time } = &event.kind {
            fragment.push_str(" seek=\"");
            fragment.push_str(seek_time.as_str());
            fragment.push('"');
        }
        fragment.push_str("/>");
    }
    fragment.push_str("</p14:showEvtLst></p:ext>");

    let updated = crate::pptx::slide_patch::insert_extension_fragment(slide.blob(), &fragment)?;
    // Self-check: the patched slide must read back through discovery.
    let probe =
        litchi_opc::BlobPart::new(slide_name.clone(), ct::PML_SLIDE.into(), updated.clone());
    let discovered = load_slide_show_events(0, &probe, &mut ShowEventLoadLimits::default())?;
    if discovered.len() != events.len() {
        return Err(invalid(
            "slide-show event storage failed read-back validation",
        ));
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
    fn scans_show_events_through_markup_compatibility() {
        let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:mc="{MCE}" xmlns:p14="{P14_NAMESPACE}" mc:Ignorable="p14"><p:extLst><p:ext uri="{SHOW_EVENT_EXTENSION_URI}"><p14:showEvtLst><p14:triggerEvt type="onClick" time="1.5s" objId="6"/><p14:seekEvt time="2000ms" objId="4" seek="10.379s"/><p14:nullEvt time="3" objId="5"/></p14:showEvtLst></p:ext></p:extLst></p:sld>"#
        );
        let events =
            scan_slide_show_events(4, xml.as_bytes(), &mut ShowEventLoadLimits::default()).unwrap();

        assert_eq!(events.len(), 3);
        assert_eq!(events[0].slide_index(), 4);
        assert_eq!(events[0].event_index(), 0);
        assert_eq!(events[0].kind(), &EventKind::Trigger(Trigger::OnClick));
        assert_eq!(events[0].time(), &Offset::parse("1.5s").unwrap());
        assert_eq!(events[0].object_id(), 6);
        assert_eq!(events[0].seek_time(), None);
        assert!(matches!(
            events[1].kind(),
            EventKind::Seek { at }
                if at == &Offset::parse("10.379s").unwrap()
        ));
        assert_eq!(
            events[1].seek_time(),
            Some(&Offset::parse("10.379s").unwrap())
        );
        assert_eq!(events[2].kind(), &EventKind::Null);
    }

    #[test]
    fn rejects_malformed_show_events() {
        let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:p14="{P14_NAMESPACE}"><p:extLst><p:ext uri="{SHOW_EVENT_EXTENSION_URI}"><p14:showEvtLst><p14:seekEvt time="1..2s" objId="4" seek="0"/></p14:showEvtLst></p:ext></p:extLst></p:sld>"#
        );

        assert!(
            scan_slide_show_events(0, xml.as_bytes(), &mut ShowEventLoadLimits::default()).is_err()
        );
    }

    #[test]
    fn retains_office_none_trigger_and_rejects_non_domain_bookmark_trigger() {
        let accepted = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:p14="{P14_NAMESPACE}"><p:extLst><p:ext uri="{SHOW_EVENT_EXTENSION_URI}"><p14:showEvtLst><p14:triggerEvt type="none" time="0" objId="1"/></p14:showEvtLst></p:ext></p:extLst></p:sld>"#
        );
        let events =
            scan_slide_show_events(0, accepted.as_bytes(), &mut ShowEventLoadLimits::default())
                .unwrap();
        assert_eq!(events[0].kind(), &EventKind::Trigger(Trigger::None));

        let rejected = accepted.replace("type=\"none\"", "type=\"onMediaBookmark\"");
        assert!(
            scan_slide_show_events(0, rejected.as_bytes(), &mut ShowEventLoadLimits::default(),)
                .is_err()
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

    fn sample_events() -> Vec<EventDraft> {
        vec![
            EventDraft::trigger(Trigger::OnClick, Offset::parse("1.5s").unwrap(), 6),
            EventDraft::play(Offset::secs(2), 4),
            EventDraft::seek(Offset::ms(2500), 4, Offset::parse("10.379s").unwrap()),
            EventDraft::null(Offset::ms(3), 5),
        ]
    }

    #[test]
    fn stores_show_events_and_discovers_them_round_trip() {
        let (mut package, slide_name) = slide_package("");
        let events = sample_events();
        store_slide_show_events(&mut package, &slide_name, &events).unwrap();

        let slide = package.get_part(&slide_name).unwrap();
        let discovered =
            load_slide_show_events(0, slide, &mut ShowEventLoadLimits::default()).unwrap();
        assert_eq!(discovered.len(), events.len());
        assert_eq!(discovered[0].kind(), &EventKind::Trigger(Trigger::OnClick));
        assert_eq!(discovered[0].time(), &Offset::ms(1500));
        assert_eq!(discovered[0].object_id(), 6);
        assert_eq!(discovered[1].kind(), &EventKind::Play);
        assert!(matches!(
            discovered[2].kind(),
            EventKind::Seek { at }
                if at == &Offset::parse("10.379s").unwrap()
        ));
        assert_eq!(
            discovered[2].seek_time(),
            Some(&Offset::parse("10.379s").unwrap())
        );
        assert_eq!(discovered[3].kind(), &EventKind::Null);

        // A second storage on the same slide is rejected (no replacement).
        assert!(store_slide_show_events(&mut package, &slide_name, &events).is_err());
    }

    #[test]
    fn stores_show_events_into_existing_extension_list() {
        let (mut package, slide_name) = slide_package(
            r#"<p:extLst><p:ext uri="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"/></p:extLst>"#,
        );
        store_slide_show_events(&mut package, &slide_name, &sample_events()).unwrap();
        let slide = package.get_part(&slide_name).unwrap();
        let xml = String::from_utf8(slide.blob().to_vec()).unwrap();
        assert!(xml.contains("{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"));
        assert!(xml.contains("p14:showEvtLst"));
        assert_eq!(
            load_slide_show_events(0, slide, &mut ShowEventLoadLimits::default())
                .unwrap()
                .len(),
            4
        );
    }

    #[test]
    fn rejects_invalid_show_event_storage() {
        let (mut package, slide_name) = slide_package("");
        // No events.
        assert!(store_slide_show_events(&mut package, &slide_name, &[]).is_err());
        // Seek owns its offset in the action variant; non-seek actions cannot
        // carry one in memory or emit one during serialization.
        let seek = EventDraft::seek(Offset::secs(1), 1, Offset::ms(250));
        assert!(matches!(
            seek.kind(),
            EventKind::Seek { at } if at == &Offset::ms(250)
        ));
        assert_eq!(EventDraft::play(Offset::secs(1), 1).seek_time(), None);
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
        assert!(store_slide_show_events(&mut package, &wrong, &sample_events()).is_err());
        // Rejection leaves the slide without an extension list.
        let slide = package.get_part(&slide_name).unwrap();
        assert!(
            load_slide_show_events(0, slide, &mut ShowEventLoadLimits::default())
                .unwrap()
                .is_empty()
        );
    }
}
