//! Bounded, inert PowerPoint slide-show event discovery.
//!
//! Slide-show events are retained as persisted document history only. This
//! module never replays, renders, seeks, pauses, resumes, stops, or otherwise
//! executes a recorded action.

use crate::common::xml::unqualified_attribute_value;
use crate::common::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::is_presentationml_name;
use litchi_opc::Part;
use litchi_opc::constants::content_type as ct;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
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
const MAX_TIME_OFFSET_BYTES: usize = 64;

/// A trigger type recorded by a PowerPoint slide show.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PptxSlideShowTrigger {
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
    OnMediaBookmark,
}

/// The recorded action represented by a PowerPoint slide-show event.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PptxSlideShowEventKind {
    Trigger(PptxSlideShowTrigger),
    Play,
    Stop,
    Pause,
    Resume,
    Seek,
    /// A reserved unknown event record for future PowerPoint extensions.
    Null,
}

/// A bounded, inert event record persisted for a PowerPoint slide show.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PptxSlideShowEvent {
    slide_index: usize,
    event_index: usize,
    kind: PptxSlideShowEventKind,
    time: String,
    object_id: u32,
    seek_time: Option<String>,
}

impl PptxSlideShowEvent {
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
    pub fn kind(&self) -> PptxSlideShowEventKind {
        self.kind
    }

    /// Return the stored universal time offset in the slide timeline.
    #[inline]
    pub fn time(&self) -> &str {
        &self.time
    }

    /// Return the DrawingML object identifier targeted by this event.
    #[inline]
    pub fn object_id(&self) -> u32 {
        self.object_id
    }

    /// Return the stored media-stream offset for a seek event.
    #[inline]
    pub fn seek_time(&self) -> Option<&str> {
        self.seek_time.as_deref()
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
) -> Result<Vec<PptxSlideShowEvent>> {
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
) -> Result<Vec<PptxSlideShowEvent>> {
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
                    slide_index,
                    &mut events,
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
                    slide_index,
                    &mut events,
                    limits,
                )?;
                if kind == ElementKind::Root {
                    saw_root = true;
                    closed_root = true;
                }
            },
            Event::End(element) => {
                let kind = stack
                    .pop()
                    .ok_or_else(|| invalid("invalid slide XML nesting"))?;
                finish_element(kind, &namespace, element.name())?;
                if kind == ElementKind::Root {
                    closed_root = true;
                }
            },
            Event::Text(text) => {
                if stack.last().copied().is_some_and(ElementKind::is_known)
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid("slide-show event markup cannot contain text"));
                }
            },
            Event::CData(text) => {
                if stack.last().copied().is_some_and(ElementKind::is_known)
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid("slide-show event markup cannot contain text"));
                }
            },
            Event::GeneralRef(_) if stack.last().copied().is_some_and(ElementKind::is_known) => {
                return Err(invalid(
                    "slide-show event markup cannot contain entity references",
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
    events: &mut Vec<PptxSlideShowEvent>,
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
        return Ok(empty.then_some(ElementKind::Other).unwrap_or(kind));
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
) -> Result<PptxSlideShowEvent> {
    let time = required_attribute(element, b"time", decoder)?;
    validate_time_offset(&time)?;
    let object_id = parse_object_id(required_attribute(element, b"objId", decoder)?)?;
    let (kind, seek_time) = match element.local_name().as_ref() {
        b"triggerEvt" => (
            PptxSlideShowEventKind::Trigger(parse_trigger(required_attribute(
                element, b"type", decoder,
            )?)?),
            None,
        ),
        b"playEvt" => (PptxSlideShowEventKind::Play, None),
        b"stopEvt" => (PptxSlideShowEventKind::Stop, None),
        b"pauseEvt" => (PptxSlideShowEventKind::Pause, None),
        b"resumeEvt" => (PptxSlideShowEventKind::Resume, None),
        b"seekEvt" => {
            let seek_time = required_attribute(element, b"seek", decoder)?;
            validate_time_offset(&seek_time)?;
            (PptxSlideShowEventKind::Seek, Some(seek_time))
        },
        b"nullEvt" => (PptxSlideShowEventKind::Null, None),
        _ => return Err(invalid("unsupported slide-show event element")),
    };
    Ok(PptxSlideShowEvent {
        slide_index,
        event_index,
        kind,
        time,
        object_id,
        seek_time,
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

fn parse_trigger(value: String) -> Result<PptxSlideShowTrigger> {
    match value.as_str() {
        "none" => Ok(PptxSlideShowTrigger::None),
        "onBegin" => Ok(PptxSlideShowTrigger::OnBegin),
        "onEnd" => Ok(PptxSlideShowTrigger::OnEnd),
        "begin" => Ok(PptxSlideShowTrigger::Begin),
        "end" => Ok(PptxSlideShowTrigger::End),
        "onClick" => Ok(PptxSlideShowTrigger::OnClick),
        "onDblClick" => Ok(PptxSlideShowTrigger::OnDoubleClick),
        "onMouseOver" => Ok(PptxSlideShowTrigger::OnMouseOver),
        "onMouseOut" => Ok(PptxSlideShowTrigger::OnMouseOut),
        "onNext" => Ok(PptxSlideShowTrigger::OnNext),
        "onPrev" => Ok(PptxSlideShowTrigger::OnPrevious),
        "onStopAudio" => Ok(PptxSlideShowTrigger::OnStopAudio),
        "onMediaBookmark" => Ok(PptxSlideShowTrigger::OnMediaBookmark),
        _ => Err(invalid(format!(
            "unsupported slide-show trigger type '{value}'"
        ))),
    }
}

fn validate_time_offset(value: &str) -> Result<()> {
    if value.len() > MAX_TIME_OFFSET_BYTES {
        return Err(limit("slide-show event time offset bytes"));
    }
    let split = value
        .find(|character: char| !character.is_ascii_digit() && character != '.')
        .unwrap_or(value.len());
    let (number, unit) = value.split_at(split);
    let mut pieces = number.split('.');
    let whole = pieces.next().unwrap_or_default();
    let fraction = pieces.next();
    if whole.is_empty()
        || !whole.bytes().all(|byte| byte.is_ascii_digit())
        || fraction
            .is_some_and(|part| part.is_empty() || !part.bytes().all(|byte| byte.is_ascii_digit()))
        || pieces.next().is_some()
        || !matches!(unit, "" | "h" | "min" | "s" | "ms" | "µs" | "ns")
    {
        return Err(invalid(format!(
            "invalid slide-show event universal time offset '{value}'"
        )));
    }
    Ok(())
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
        assert_eq!(
            events[0].kind(),
            PptxSlideShowEventKind::Trigger(PptxSlideShowTrigger::OnClick)
        );
        assert_eq!(events[0].time(), "1.5s");
        assert_eq!(events[0].object_id(), 6);
        assert_eq!(events[0].seek_time(), None);
        assert_eq!(events[1].kind(), PptxSlideShowEventKind::Seek);
        assert_eq!(events[1].seek_time(), Some("10.379s"));
        assert_eq!(events[2].kind(), PptxSlideShowEventKind::Null);
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
}
