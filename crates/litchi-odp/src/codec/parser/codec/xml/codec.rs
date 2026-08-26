//! XML token and event traversal for the ODP parser.

use super::super::{
    Action, AnimationKind, AnimationNode, BytesRef, BytesStart, DRAW_NAMESPACE, DrawingHyperlink,
    DrawingShapeKind, Effect, EffectDirection, Element, EnhancedGeometry, EnhancedGeometryChild,
    EnhancedGeometryChildKind, Error, Event, EventListener, Kind, Node, NsClass, NsReader,
    PRESENTATION_NAMESPACE, ParagraphText, Parser, Result, SCRIPT_NAMESPACE, STYLE_NAMESPACE,
    ScriptEventListener, Shape, ShapeBuilder, ShapeContainerScope, ShapeEventListener, ShapeType,
    Slide, Speed, Transition, TransitionStyleDefinition, TransitionStyles, XLINK_NAMESPACE,
    XmlNamespace, XmlVersion, validate_legacy_animation_root,
};
use super::validation::ElementAttrs;
use litchi_core::{SequentialTextWriter, TextObjectKind, TextOutputError};
use quick_xml::name::ResolveResult;
use std::io::Write;

impl Parser {
    pub(super) fn parse_animation_node(
        reader: &mut NsReader<&[u8]>,
        collector: &mut TransitionStyleCollector,
        start: &BytesStart<'_>,
        kind: Kind,
        depth: usize,
        node_count: &mut usize,
    ) -> Result<Node> {
        if depth > 128 {
            return Err(Error::InvalidFormat(
                "ODP animation nesting exceeds 128 levels".to_string(),
            ));
        }
        *node_count = node_count
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("ODP animation node count overflow".to_string()))?;
        if *node_count > 65_536 {
            return Err(Error::InvalidFormat(
                "ODP animation tree exceeds 65536 nodes".to_string(),
            ));
        }
        let attributes = Self::animation_attributes(reader, start)?;
        let mut children = Vec::new();
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            let ns_class = NsClass::from_resolve(&namespace);
            collector.feed(reader, ns_class, &event);
            match event {
                Event::Start(ref child) | Event::Empty(ref child) => {
                    if ns_class != NsClass::Animation {
                        return Err(Error::InvalidFormat(format!(
                            "anim:{} contains a non-animation element",
                            kind.local_name()
                        )));
                    }
                    let Some(child_kind) = Kind::from_local_name(child.local_name().as_ref())
                    else {
                        return Err(Error::InvalidFormat(format!(
                            "unknown ODF animation element '{}'",
                            String::from_utf8_lossy(child.local_name().as_ref())
                        )));
                    };
                    if !kind.allows_child(child_kind) {
                        return Err(Error::InvalidFormat(format!(
                            "anim:{} cannot contain anim:{}",
                            kind.local_name(),
                            child_kind.local_name()
                        )));
                    }
                    let node = if matches!(event, Event::Empty(_)) {
                        *node_count = node_count.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat("ODP animation node count overflow".to_string())
                        })?;
                        if *node_count > 65_536 {
                            return Err(Error::InvalidFormat(
                                "ODP animation tree exceeds 65536 nodes".to_string(),
                            ));
                        }
                        Node::from_parsed(
                            child_kind,
                            Self::animation_attributes(reader, child)?,
                            Vec::new(),
                        )
                    } else {
                        Self::parse_animation_node(
                            reader,
                            collector,
                            child,
                            child_kind,
                            depth + 1,
                            node_count,
                        )?
                    };
                    children.push(node);
                },
                Event::End(ref end) => {
                    if ns_class != NsClass::Animation
                        || end.local_name().as_ref() != kind.local_name().as_bytes()
                    {
                        return Err(Error::InvalidFormat(format!(
                            "unexpected closing element in anim:{}",
                            kind.local_name()
                        )));
                    }
                    return Ok(Node::from_parsed(kind, attributes, children));
                },
                Event::Text(ref text) => {
                    let decoded = Self::decode_text(text)?;
                    if !decoded.trim().is_empty() {
                        return Err(Error::InvalidFormat(format!(
                            "anim:{} cannot contain text",
                            kind.local_name()
                        )));
                    }
                },
                Event::CData(ref text) => {
                    let content = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid animation CDATA: {error}"))
                    })?;
                    if !content.trim().is_empty() {
                        return Err(Error::InvalidFormat(format!(
                            "anim:{} cannot contain text",
                            kind.local_name()
                        )));
                    }
                },
                Event::Eof => {
                    return Err(Error::InvalidFormat(format!(
                        "unterminated anim:{} element",
                        kind.local_name()
                    )));
                },
                Event::GeneralRef(_) => {
                    return Err(Error::InvalidFormat(format!(
                        "anim:{} cannot contain character references",
                        kind.local_name()
                    )));
                },
                Event::Comment(_) | Event::Decl(_) | Event::PI(_) | Event::DocType(_) => {},
            }
            buffer.clear();
        }
    }

    pub(super) fn parse_legacy_animation_node(
        reader: &mut NsReader<&[u8]>,
        collector: &mut TransitionStyleCollector,
        start: &BytesStart<'_>,
        kind: AnimationKind,
        depth: usize,
        node_count: &mut usize,
    ) -> Result<AnimationNode> {
        if depth > 128 {
            return Err(Error::InvalidFormat(
                "legacy ODP animation nesting exceeds 128 levels".to_string(),
            ));
        }
        *node_count = node_count.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("legacy ODP animation node count overflow".to_string())
        })?;
        if *node_count > 65_536 {
            return Err(Error::InvalidFormat(
                "legacy ODP animation tree exceeds 65536 nodes".to_string(),
            ));
        }
        let attributes = Self::animation_attributes(reader, start)?;
        let mut children = Vec::new();
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            let ns_class = NsClass::from_resolve(&namespace);
            collector.feed(reader, ns_class, &event);
            match event {
                Event::Start(ref child) | Event::Empty(ref child) => {
                    if ns_class != NsClass::Presentation {
                        return Err(Error::InvalidFormat(format!(
                            "presentation:{} contains a foreign element",
                            kind.local_name()
                        )));
                    }
                    let child_kind = AnimationKind::from_local_name(child.local_name().as_ref())
                        .ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "unknown legacy presentation animation element '{}'",
                                String::from_utf8_lossy(child.local_name().as_ref())
                            ))
                        })?;
                    if !kind.allows_child(child_kind) {
                        return Err(Error::InvalidFormat(format!(
                            "presentation:{} cannot contain presentation:{}",
                            kind.local_name(),
                            child_kind.local_name()
                        )));
                    }
                    let node = if matches!(event, Event::Empty(_)) {
                        *node_count = node_count.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(
                                "legacy ODP animation node count overflow".to_string(),
                            )
                        })?;
                        if *node_count > 65_536 {
                            return Err(Error::InvalidFormat(
                                "legacy ODP animation tree exceeds 65536 nodes".to_string(),
                            ));
                        }
                        AnimationNode::from_parsed(
                            child_kind,
                            Self::animation_attributes(reader, child)?,
                            Vec::new(),
                        )
                    } else {
                        Self::parse_legacy_animation_node(
                            reader,
                            collector,
                            child,
                            child_kind,
                            depth + 1,
                            node_count,
                        )?
                    };
                    children.push(node);
                },
                Event::End(ref end)
                    if ns_class == NsClass::Presentation
                        && end.local_name().as_ref() == kind.local_name().as_bytes() =>
                {
                    return Ok(AnimationNode::from_parsed(kind, attributes, children));
                },
                Event::Text(ref text) if !Self::decode_text(text)?.trim().is_empty() => {
                    return Err(Error::InvalidFormat(
                        "legacy presentation animations cannot contain text".to_string(),
                    ));
                },
                Event::CData(ref text)
                    if !text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::InvalidFormat(error.to_string()))?
                        .trim()
                        .is_empty() =>
                {
                    return Err(Error::InvalidFormat(
                        "legacy presentation animations cannot contain text".to_string(),
                    ));
                },
                Event::Eof => {
                    return Err(Error::InvalidFormat(
                        "unterminated legacy presentation animation tree".to_string(),
                    ));
                },
                Event::End(_) | Event::GeneralRef(_) => {
                    return Err(Error::InvalidFormat(
                        "invalid content in legacy presentation animation tree".to_string(),
                    ));
                },
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_) => {},
            }
            buffer.clear();
        }
    }

    pub(super) fn parse_enhanced_geometry(
        reader: &mut NsReader<&[u8]>,
        collector: &mut TransitionStyleCollector,
        element: &BytesStart<'_>,
    ) -> Result<EnhancedGeometry> {
        let attributes = Self::exact_geometry_attributes(reader, element)?;
        let mut children = Vec::new();
        let mut handle_seen = false;
        let mut buffer = Vec::new();
        loop {
            // Tokenization failures use the same mapping as every other read
            // site in this module: historically the transition-definitions
            // pre-scan tokenized the whole content part before this parser
            // could run, so this branch could only ever surface through the
            // pre-scan's "XML parsing error" message.
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            let ns_class = NsClass::from_resolve(&namespace);
            collector.feed(reader, ns_class, &event);
            match event {
                Event::Start(ref child) | Event::Empty(ref child)
                    if ns_class == NsClass::Drawing
                        && matches!(child.local_name().as_ref(), b"equation" | b"handle") =>
                {
                    if children.len() >= 65_536 {
                        return Err(Error::InvalidFormat(
                            "enhanced geometry exceeds 65536 equations and handles".to_string(),
                        ));
                    }
                    let kind = if child.local_name().as_ref() == b"equation" {
                        if handle_seen {
                            return Err(Error::InvalidFormat(
                                "draw:equation cannot follow draw:handle".to_string(),
                            ));
                        }
                        EnhancedGeometryChildKind::Equation
                    } else {
                        handle_seen = true;
                        EnhancedGeometryChildKind::Handle
                    };
                    children.push(EnhancedGeometryChild {
                        kind,
                        attributes: Self::exact_geometry_attributes(reader, child)?,
                    });
                    if matches!(event, Event::Start(_)) {
                        Self::consume_empty_content(
                            reader,
                            collector,
                            DRAW_NAMESPACE,
                            child.local_name().as_ref(),
                            kind.element_name(),
                        )?;
                    }
                },
                Event::End(ref end)
                    if ns_class == NsClass::Drawing
                        && end.local_name().as_ref() == b"enhanced-geometry" =>
                {
                    return Ok(EnhancedGeometry {
                        attributes,
                        children,
                    });
                },
                Event::Text(ref text) if Self::decode_text(text)?.trim().is_empty() => {},
                Event::CData(ref text)
                    if text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::InvalidFormat(error.to_string()))?
                        .trim()
                        .is_empty() => {},
                Event::Comment(_) | Event::PI(_) => {},
                Event::Eof => {
                    return Err(Error::InvalidFormat(
                        "unterminated draw:enhanced-geometry".to_string(),
                    ));
                },
                Event::Start(_)
                | Event::Empty(_)
                | Event::End(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Decl(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => {
                    return Err(Error::InvalidFormat(
                        "draw:enhanced-geometry may only contain equations and handles".to_string(),
                    ));
                },
            }
            buffer.clear();
        }
    }

    pub(super) fn script_event_listener(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<ScriptEventListener> {
        let mut attributes = ElementAttrs::new(element);
        let event_name =
            attributes.required(reader, SCRIPT_NAMESPACE, b"event-name", "script:event-name")?;
        let language =
            attributes.required(reader, SCRIPT_NAMESPACE, b"language", "script:language")?;
        let macro_name = attributes.get(reader, SCRIPT_NAMESPACE, b"macro-name")?;
        let href = attributes.get(reader, XLINK_NAMESPACE, b"href")?;
        let link_type = attributes.get(reader, XLINK_NAMESPACE, b"type")?;
        if href.is_some() {
            attributes.require_simple_xlink(reader, "script:event-listener")?;
        } else if link_type.is_some() {
            return Err(Error::InvalidFormat(
                "script:event-listener xlink:type requires xlink:href".to_string(),
            ));
        }
        let listener = ScriptEventListener {
            event_name,
            language,
            macro_name,
            href,
            actuate_on_request: Self::parse_on_request(
                attributes
                    .get(reader, XLINK_NAMESPACE, b"actuate")?
                    .as_deref(),
                "script:event-listener",
            )?,
        };
        listener.validate()?;
        Ok(listener)
    }

    pub(super) fn presentation_event_listener(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<EventListener> {
        let mut attributes = ElementAttrs::new(element);
        let event_name =
            attributes.required(reader, SCRIPT_NAMESPACE, b"event-name", "script:event-name")?;
        let action = Action::parse(&attributes.required(
            reader,
            PRESENTATION_NAMESPACE,
            b"action",
            "presentation:action",
        )?)?;
        let mut listener = EventListener::new(event_name, action)?;
        listener.effect = attributes
            .get(reader, PRESENTATION_NAMESPACE, b"effect")?
            .map(Effect::new)
            .transpose()?;
        listener.direction = attributes
            .get(reader, PRESENTATION_NAMESPACE, b"direction")?
            .map(EffectDirection::new)
            .transpose()?;
        listener.speed = attributes
            .get(reader, PRESENTATION_NAMESPACE, b"speed")?
            .map(|value| Speed::parse(&value))
            .transpose()?;
        listener.start_scale = attributes.get(reader, PRESENTATION_NAMESPACE, b"start-scale")?;
        listener.href = attributes.get(reader, XLINK_NAMESPACE, b"href")?;
        let link_type = attributes.get(reader, XLINK_NAMESPACE, b"type")?;
        if listener.href.is_some() {
            attributes.require_simple_xlink(reader, "presentation:event-listener")?;
        } else if link_type.is_some() {
            return Err(Error::InvalidFormat(
                "presentation:event-listener xlink:type requires xlink:href".to_string(),
            ));
        }
        listener.show_embed = match attributes.get(reader, XLINK_NAMESPACE, b"show")?.as_deref() {
            None => false,
            Some("embed") => true,
            Some(value) => {
                return Err(Error::InvalidFormat(format!(
                    "invalid presentation:event-listener xlink:show '{value}'"
                )));
            },
        };
        listener.actuate_on_request = Self::parse_on_request(
            attributes
                .get(reader, XLINK_NAMESPACE, b"actuate")?
                .as_deref(),
            "presentation:event-listener",
        )?;
        listener.verb = attributes
            .get(reader, PRESENTATION_NAMESPACE, b"verb")?
            .map(|value| {
                value.parse::<u64>().map_err(|_err| {
                    Error::InvalidFormat(format!("invalid presentation:verb '{value}'"))
                })
            })
            .transpose()?;
        listener.validate()?;
        Ok(listener)
    }

    pub(super) fn consume_empty_content(
        reader: &mut NsReader<&[u8]>,
        collector: &mut TransitionStyleCollector,
        namespace_uri: &[u8],
        local_name: &[u8],
        description: &str,
    ) -> Result<()> {
        let mut buffer = Vec::new();
        let expected_class = NsClass::from_uri(namespace_uri);
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            let ns_class = NsClass::from_resolve(&namespace);
            collector.feed(reader, ns_class, &event);
            match event {
                Event::End(ref end)
                    if expected_class != NsClass::Other
                        && ns_class == expected_class
                        && end.local_name().as_ref() == local_name =>
                {
                    return Ok(());
                },
                Event::Text(ref text) if Self::decode_text(text)?.trim().is_empty() => {},
                Event::CData(ref text)
                    if text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::InvalidFormat(error.to_string()))?
                        .trim()
                        .is_empty() => {},
                Event::Eof => {
                    return Err(Error::InvalidFormat(format!(
                        "unterminated {description} element"
                    )));
                },
                Event::Start(_)
                | Event::Empty(_)
                | Event::End(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => {
                    return Err(Error::InvalidFormat(format!(
                        "{description} must not contain content"
                    )));
                },
            }
            buffer.clear();
        }
    }

    pub(super) fn parse_listener_body(
        reader: &mut NsReader<&[u8]>,
        collector: &mut TransitionStyleCollector,
        mut listener: EventListener,
    ) -> Result<EventListener> {
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            let ns_class = NsClass::from_resolve(&namespace);
            collector.feed(reader, ns_class, &event);
            match event {
                Event::Start(ref element) | Event::Empty(ref element)
                    if ns_class == NsClass::Presentation
                        && element.local_name().as_ref() == b"sound" =>
                {
                    if listener.sound.is_some() {
                        return Err(Error::InvalidFormat(
                            "presentation event listener contains multiple sounds".to_string(),
                        ));
                    }
                    Self::require_simple_xlink(reader, element, "presentation:sound")?;
                    listener.sound = Some(Self::parse_transition_sound(reader, element)?);
                    if matches!(event, Event::Start(_)) {
                        Self::consume_empty_content(
                            reader,
                            collector,
                            PRESENTATION_NAMESPACE,
                            b"sound",
                            "presentation:sound",
                        )?;
                    }
                },
                Event::End(ref end)
                    if ns_class == NsClass::Presentation
                        && end.local_name().as_ref() == b"event-listener" =>
                {
                    listener.validate()?;
                    return Ok(listener);
                },
                Event::Text(ref text) if Self::decode_text(text)?.trim().is_empty() => {},
                Event::CData(ref text)
                    if text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::InvalidFormat(error.to_string()))?
                        .trim()
                        .is_empty() => {},
                Event::Eof => {
                    return Err(Error::InvalidFormat(
                        "unterminated presentation:event-listener".to_string(),
                    ));
                },
                Event::Start(_)
                | Event::Empty(_)
                | Event::End(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => {
                    return Err(Error::InvalidFormat(
                        "presentation:event-listener may only contain presentation:sound"
                            .to_string(),
                    ));
                },
            }
            buffer.clear();
        }
    }

    pub(super) fn parse_event_listeners(
        reader: &mut NsReader<&[u8]>,
        collector: &mut TransitionStyleCollector,
    ) -> Result<Vec<ShapeEventListener>> {
        let mut listeners = Vec::new();
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            let ns_class = NsClass::from_resolve(&namespace);
            collector.feed(reader, ns_class, &event);
            match event {
                Event::Start(ref element) | Event::Empty(ref element)
                    if ns_class == NsClass::Script
                        && element.local_name().as_ref() == b"event-listener" =>
                {
                    if listeners.len() >= 4096 {
                        return Err(Error::InvalidFormat(
                            "ODP shape exceeds 4096 event listeners".to_string(),
                        ));
                    }
                    let listener = Self::script_event_listener(reader, element)?;
                    if matches!(event, Event::Start(_)) {
                        Self::consume_empty_content(
                            reader,
                            collector,
                            SCRIPT_NAMESPACE,
                            b"event-listener",
                            "script:event-listener",
                        )?;
                    }
                    listeners.push(ShapeEventListener::Script(listener));
                },
                Event::Start(ref element) | Event::Empty(ref element)
                    if ns_class == NsClass::Presentation
                        && element.local_name().as_ref() == b"event-listener" =>
                {
                    if listeners.len() >= 4096 {
                        return Err(Error::InvalidFormat(
                            "ODP shape exceeds 4096 event listeners".to_string(),
                        ));
                    }
                    let listener = Self::presentation_event_listener(reader, element)?;
                    let parsed_listener = if matches!(event, Event::Start(_)) {
                        Self::parse_listener_body(reader, collector, listener)?
                    } else {
                        listener
                    };
                    listeners.push(ShapeEventListener::Action(Box::new(parsed_listener)));
                },
                Event::End(ref end)
                    if ns_class == NsClass::Office
                        && end.local_name().as_ref() == b"event-listeners" =>
                {
                    return Ok(listeners);
                },
                Event::Text(ref text) if Self::decode_text(text)?.trim().is_empty() => {},
                Event::CData(ref text)
                    if text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::InvalidFormat(error.to_string()))?
                        .trim()
                        .is_empty() => {},
                Event::Eof => {
                    return Err(Error::InvalidFormat(
                        "unterminated office:event-listeners".to_string(),
                    ));
                },
                Event::Start(_)
                | Event::Empty(_)
                | Event::End(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => {
                    return Err(Error::InvalidFormat(
                        "office:event-listeners may only contain script or presentation listeners"
                            .to_string(),
                    ));
                },
            }
            buffer.clear();
        }
    }

    pub(super) fn decode_text(text: &quick_xml::events::BytesText<'_>) -> Result<String> {
        let decoded = text
            .xml_content(XmlVersion::Explicit1_0)
            .map_err(|error| Error::InvalidFormat(format!("invalid presentation text: {error}")))?;
        Ok(decoded.into_owned())
    }

    pub(super) fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
        if let Some(character) = reference.resolve_char_ref().map_err(|error| {
            Error::InvalidFormat(format!("invalid presentation character reference: {error}"))
        })? {
            return Ok(character.to_string());
        }
        let name = reference.decode().map_err(|error| {
            Error::InvalidFormat(format!("invalid presentation entity reference: {error}"))
        })?;
        match name.as_ref() {
            "amp" => Ok("&".to_string()),
            "lt" => Ok("<".to_string()),
            "gt" => Ok(">".to_string()),
            "quot" => Ok("\"".to_string()),
            "apos" => Ok("'".to_string()),
            _ => Err(Error::InvalidFormat(format!(
                "unsupported presentation entity reference '&{name};'"
            ))),
        }
    }

    pub(super) fn parse_transition_style_definitions(xml: &str) -> Result<TransitionStyles> {
        let mut reader = NsReader::from_str(xml);
        let mut buf = Vec::new();
        let mut result = TransitionStyles::default();
        let mut current: Option<(Option<String>, bool, TransitionStyleDefinition)> = None;
        let mut in_properties = false;

        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buf)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            match event {
                Event::Start(ref element)
                    if Self::is_namespace(&namespace, STYLE_NAMESPACE)
                        && matches!(element.local_name().as_ref(), b"style" | b"default-style") =>
                {
                    let mut attributes = ElementAttrs::new(element);
                    let family = attributes.get(&reader, STYLE_NAMESPACE, b"family")?;
                    let is_drawing_page = family.as_deref() == Some("drawing-page");
                    let name = attributes.get(&reader, STYLE_NAMESPACE, b"name")?;
                    let parent = attributes.get(&reader, STYLE_NAMESPACE, b"parent-style-name")?;
                    current = Some((
                        name,
                        is_drawing_page,
                        TransitionStyleDefinition {
                            parent,
                            transition: Transition::new(),
                        },
                    ));
                },
                Event::Empty(ref element)
                    if Self::is_namespace(&namespace, STYLE_NAMESPACE)
                        && matches!(element.local_name().as_ref(), b"style" | b"default-style") =>
                {
                    let mut attributes = ElementAttrs::new(element);
                    let family = attributes.get(&reader, STYLE_NAMESPACE, b"family")?;
                    if family.as_deref() == Some("drawing-page") {
                        let name = attributes.get(&reader, STYLE_NAMESPACE, b"name")?;
                        let definition = TransitionStyleDefinition {
                            parent: attributes.get(
                                &reader,
                                STYLE_NAMESPACE,
                                b"parent-style-name",
                            )?,
                            transition: Transition::new(),
                        };
                        if let Some(style_name) = name {
                            result.named.insert(style_name, definition);
                        } else {
                            result.default = definition.transition;
                        }
                    }
                },
                Event::Start(ref element) | Event::Empty(ref element)
                    if current.as_ref().is_some_and(|(_, family, _)| *family)
                        && Self::is_namespace(&namespace, STYLE_NAMESPACE)
                        && element.local_name().as_ref() == b"drawing-page-properties" =>
                {
                    if let Some((_, _, definition)) = current.as_mut() {
                        Self::parse_transition_properties(
                            &reader,
                            element,
                            &mut definition.transition,
                        )?;
                    }
                    in_properties = matches!(event, Event::Start(_));
                },
                Event::Start(ref element) | Event::Empty(ref element)
                    if in_properties
                        && Self::is_namespace(&namespace, PRESENTATION_NAMESPACE)
                        && element.local_name().as_ref() == b"sound" =>
                {
                    if let Some((_, _, definition)) = current.as_mut() {
                        definition.transition.sound =
                            Some(Self::parse_transition_sound(&reader, element)?);
                    }
                },
                Event::End(ref element)
                    if Self::is_namespace(&namespace, STYLE_NAMESPACE)
                        && element.local_name().as_ref() == b"drawing-page-properties" =>
                {
                    in_properties = false;
                },
                Event::End(ref element)
                    if Self::is_namespace(&namespace, STYLE_NAMESPACE)
                        && matches!(element.local_name().as_ref(), b"style" | b"default-style") =>
                {
                    if let Some((name, is_drawing_page, definition)) = current.take()
                        && is_drawing_page
                    {
                        if let Some(style_name) = name {
                            result.named.insert(style_name, definition);
                        } else {
                            result.default = definition.transition;
                        }
                    }
                    in_properties = false;
                },
                Event::Eof => break,
                Event::Start(_)
                | Event::Empty(_)
                | Event::End(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => {},
            }
            buf.clear();
        }
        Ok(result)
    }

    #[cfg(test)]
    pub(crate) fn parse_slides(xml_content: &str) -> Result<Vec<Slide>> {
        Self::parse_slides_with_styles(xml_content, None)
    }

    /// Parse slides and resolve drawing-page transition styles.
    pub(crate) fn parse_slides_with_styles(
        xml_content: &str,
        styles_xml: Option<&str>,
    ) -> Result<Vec<Slide>> {
        Self::parse_pages_with_styles::<false>(
            xml_content,
            styles_xml,
            0,
            false,
            ShapeContainerScope::DrawPages,
        )
    }

    /// Parse one slide while retaining full-document validation semantics.
    pub(crate) fn parse_slide_with_styles_at(
        xml_content: &str,
        styles_xml: Option<&str>,
        index: usize,
    ) -> Result<Option<Slide>> {
        let mut slides = Self::parse_pages_with_styles::<true>(
            xml_content,
            styles_xml,
            index,
            false,
            ShapeContainerScope::DrawPages,
        )?;
        Ok(slides.pop())
    }

    /// Parse drawing pages while retaining title and text-box frames as shapes.
    #[allow(dead_code, reason = "reserved for the dedicated ODG facade")]
    pub(crate) fn parse_drawing_pages(
        xml_content: &str,
        styles_xml: Option<&str>,
    ) -> Result<Vec<Slide>> {
        Self::parse_pages_with_styles::<false>(
            xml_content,
            styles_xml,
            0,
            true,
            ShapeContainerScope::DrawPages,
        )
    }

    /// Parse `table:shapes` drawing shapes from spreadsheet content.
    ///
    /// Returns one shape list per top-level `table:table` element, in
    /// document order, retaining text-box frames as shapes. Shapes anchored
    /// inside individual table cells are not collected.
    #[allow(dead_code, reason = "reserved for the dedicated ODS facade")]
    pub(crate) fn parse_sheet_shape_tables(xml_content: &str) -> Result<Vec<Vec<Shape>>> {
        let tables = Self::parse_pages_with_styles::<false>(
            xml_content,
            None,
            0,
            true,
            ShapeContainerScope::SpreadsheetTables,
        )?;
        Ok(tables.into_iter().map(|table| table.shapes).collect())
    }

    /// Whether a fused-pass error came from tokenizing the XML stream rather
    /// than from semantic validation.
    ///
    /// Every `read_resolved_event_into` site in this module maps failures with
    /// the `XML parsing error: ` prefix and no semantic error uses it, so the
    /// prefix reliably identifies read errors. Tokenization errors always
    /// surface in transition-scan position, regardless of which parser
    /// observed them.
    fn is_xml_read_error(error: &Error) -> bool {
        matches!(error, Error::InvalidFormat(message) if message.starts_with("XML parsing error: "))
    }

    pub(super) fn parse_pages_with_styles<const SELECT_ONE: bool>(
        xml_content: &str,
        styles_xml: Option<&str>,
        selected_index: usize,
        retain_text_shapes: bool,
        container_scope: ShapeContainerScope,
    ) -> Result<Vec<Slide>> {
        let sheet_scope = container_scope == ShapeContainerScope::SpreadsheetTables;
        // Historical pass order: the styles.xml transition-definitions scan
        // completes (or fails) before anything content.xml parsing surfaces.
        let mut definitions = TransitionStyles::default();
        if let Some(styles_source) = styles_xml {
            definitions = Self::parse_transition_style_definitions(styles_source)?;
        }
        let mut collector = TransitionStyleCollector::default();
        let mut reader = NsReader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut slides = Vec::new();
        // Deferred drawing-page transition lookups, one entry per pushed
        // slide, resolved once the fused pass has collected every definition.
        let mut pending_transitions: Vec<Option<String>> = Vec::new();
        // First slide-scan error, surfaced only after transition collection
        // and inheritance resolution, matching the historical pass order.
        let mut main_error: Option<Error> = None;

        // State tracking
        let mut current_slide_text = String::new();
        let mut current_slide_title: Option<String> = None;
        let mut current_shapes: Vec<Shape> = Vec::new();
        let mut in_slide = false;
        let mut slide_index = 0;
        let mut current_notes_text = String::new();
        let mut current_notes_has_paragraph = false;
        let mut in_notes = false;
        let mut current_slide_has_segment = false;
        let mut current_page_style: Option<String> = None;
        let mut current_animations = Vec::new();
        let mut animation_node_count = 0;
        let mut current_legacy_animation = None;
        let mut legacy_animation_node_count = 0;
        let mut shape_node_count = 0usize;

        // Shape parsing state
        let mut shape_stack: Vec<ShapeBuilder> = Vec::new();
        let mut current_paragraph: Option<ParagraphText> = None;
        let mut in_media_plugin = false;
        let mut in_media_parameter = false;
        let mut current_hyperlink: Option<DrawingHyperlink> = None;
        let mut hyperlink_parent_depth = None;
        let mut hyperlink_shape_seen = false;

        // Spreadsheet `table:shapes` container state
        let mut element_depth = 0usize;
        let mut spreadsheet_depth: Option<usize> = None;
        let mut sheet_table_depth: Option<usize> = None;
        let mut sheet_shapes_depth: Option<usize> = None;
        let mut sheet_table_has_shapes = false;

        // One tokenization pass drives both logical handlers: every event is
        // fed to the transition-definition collector before the slide scan
        // reacts to it, so collection errors keep their historical precedence
        // over slide-scan errors at the same or later positions.
        let mut process = |reader: &mut NsReader<&[u8]>,
                           collector: &mut TransitionStyleCollector,
                           ns_class: NsClass,
                           event: &Event<'_>|
         -> Result<()> {
            match event {
                Event::Start(element) => {
                    element_depth = element_depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("XML element depth overflow".to_string())
                    })?;
                    let element_type = Self::classify(ns_class, element.local_name().as_ref());
                    Self::validate_three_dimensional_child_element(
                        shape_stack.last(),
                        element_type,
                    )?;
                    if in_media_parameter {
                        return Err(Error::InvalidFormat(
                            "draw:param cannot contain child elements".to_string(),
                        ));
                    }
                    if in_media_plugin && !matches!(element_type, Element::PluginParameter) {
                        return Err(Error::InvalidFormat(
                            "draw:plugin can only contain draw:param elements".to_string(),
                        ));
                    }
                    match element_type {
                        Element::Page if !sheet_scope => {
                            if in_slide {
                                if !SELECT_ONE || slide_index == selected_index {
                                    slides.push(Slide {
                                        title: current_slide_title.take(),
                                        text: std::mem::take(&mut current_slide_text),
                                        index: slide_index,
                                        notes: (!current_notes_text.is_empty())
                                            .then(|| std::mem::take(&mut current_notes_text)),
                                        transition: None,
                                        animations: std::mem::take(&mut current_animations),
                                        legacy_animation: current_legacy_animation.take(),
                                        shapes: std::mem::take(&mut current_shapes),
                                    });
                                    pending_transitions.push(current_page_style.take());
                                } else {
                                    current_slide_text.clear();
                                    current_notes_text.clear();
                                    current_animations.clear();
                                    current_legacy_animation = None;
                                    current_shapes.clear();
                                }
                                slide_index += 1;
                            }
                            current_slide_title = None;
                            current_slide_has_segment = false;
                            current_notes_has_paragraph = false;
                            current_page_style =
                                Self::get_attr(&*reader, element, DRAW_NAMESPACE, b"style-name")?;
                            in_slide = true;
                        },
                        Element::Notes if in_slide => in_notes = true,
                        Element::EnhancedGeometry if !shape_stack.is_empty() => {
                            if let Some(builder) = shape_stack.last() {
                                if builder.drawing_kind != Some(DrawingShapeKind::CustomShape) {
                                    return Err(Error::InvalidFormat(
                                        "draw:enhanced-geometry requires draw:custom-shape"
                                            .to_string(),
                                    ));
                                }
                                if builder.enhanced_geometry.is_some() {
                                    return Err(Error::InvalidFormat(
                                        "draw:custom-shape contains multiple enhanced geometries"
                                            .to_string(),
                                    ));
                                }
                            }
                            let geometry =
                                Self::parse_enhanced_geometry(reader, collector, element)?;
                            element_depth = Self::rewind_consumed_subtree(element_depth);
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.enhanced_geometry = Some(geometry);
                            }
                        },
                        Element::EnhancedGeometry
                        | Element::EnhancedEquation
                        | Element::EnhancedHandle
                            if in_slide =>
                        {
                            return Err(Error::InvalidFormat(
                                "misplaced custom-shape enhanced geometry".to_string(),
                            ));
                        },
                        Element::LegacyAnimation(kind)
                            if in_slide
                                && !in_notes
                                && shape_stack.is_empty()
                                && current_paragraph.is_none() =>
                        {
                            if kind != AnimationKind::Animations {
                                return Err(Error::InvalidFormat(
                                    "legacy presentation effects require a presentation:animations root"
                                        .to_string(),
                                ));
                            }
                            if current_legacy_animation.is_some() {
                                return Err(Error::InvalidFormat(
                                    "ODP slide contains multiple presentation:animations roots"
                                        .to_string(),
                                ));
                            }
                            let root = Self::parse_legacy_animation_node(
                                reader,
                                collector,
                                element,
                                kind,
                                1,
                                &mut legacy_animation_node_count,
                            )?;
                            element_depth = Self::rewind_consumed_subtree(element_depth);
                            validate_legacy_animation_root(&root)?;
                            current_legacy_animation = Some(root);
                        },
                        Element::Plugin if !shape_stack.is_empty() => {
                            if let Some(builder) = shape_stack.last_mut() {
                                if !builder.is_frame {
                                    return Err(Error::InvalidFormat(
                                        "draw:plugin must be contained directly by draw:frame"
                                            .to_string(),
                                    ));
                                }
                                if builder.media.is_some() {
                                    return Err(Error::InvalidFormat(
                                        "ODP frame contains multiple draw:plugin elements"
                                            .to_string(),
                                    ));
                                }
                                builder.shape_type = ShapeType::GraphicFrame;
                                builder.media = Some(Self::media_reference(&*reader, element)?);
                                in_media_plugin = true;
                            }
                        },
                        Element::Plugin if in_slide => {
                            return Err(Error::InvalidFormat(
                                "draw:plugin must be contained by a drawing shape".to_string(),
                            ));
                        },
                        Element::PluginParameter
                            if in_media_plugin
                                && !in_media_parameter
                                && !shape_stack.is_empty() =>
                        {
                            if let Some(media) = shape_stack
                                .last_mut()
                                .and_then(|builder| builder.media.as_mut())
                            {
                                media.add_parameter(Self::media_parameter(&*reader, element)?)?;
                                in_media_parameter = true;
                            }
                        },
                        Element::DrawingHyperlink
                            if in_slide && !in_notes && current_hyperlink.is_none() =>
                        {
                            current_hyperlink = Some(Self::drawing_hyperlink(&*reader, element)?);
                            hyperlink_parent_depth = Some(shape_stack.len());
                            hyperlink_shape_seen = false;
                        },
                        Element::DrawingHyperlink if in_slide => {
                            return Err(Error::InvalidFormat(
                                "nested or misplaced draw:a presentation hyperlink".to_string(),
                            ));
                        },
                        Element::EventListeners if !shape_stack.is_empty() => {
                            if let Some(builder) = shape_stack.last_mut() {
                                if builder
                                    .drawing_kind
                                    .is_some_and(DrawingShapeKind::is_three_dimensional)
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D shapes cannot contain presentation event listeners"
                                            .to_string(),
                                    ));
                                }
                                if builder.event_listeners_seen {
                                    return Err(Error::InvalidFormat(
                                        "ODP shape contains multiple office:event-listeners elements"
                                            .to_string(),
                                    ));
                                }
                                builder.event_listeners =
                                    Self::parse_event_listeners(reader, collector)?;
                                element_depth = Self::rewind_consumed_subtree(element_depth);
                                builder.event_listeners_seen = true;
                            }
                        },
                        Element::EventListeners
                        | Element::EventListener
                        | Element::ScriptEventListener
                        | Element::Sound
                            if in_slide =>
                        {
                            return Err(Error::InvalidFormat(
                                "presentation event metadata must be contained by a shape's office:event-listeners"
                                    .to_string(),
                            ));
                        },
                        _ if in_media_parameter => {
                            return Err(Error::InvalidFormat(
                                "draw:param cannot contain child elements".to_string(),
                            ));
                        },
                        _ if in_media_plugin => {
                            return Err(Error::InvalidFormat(
                                "draw:plugin can only contain draw:param elements".to_string(),
                            ));
                        },
                        Element::TextParagraph if in_slide => {
                            if current_paragraph.is_some() {
                                return Err(Error::InvalidFormat(
                                    "nested ODP text paragraphs are not supported".to_string(),
                                ));
                            }
                            current_paragraph = Some(ParagraphText::default());
                        },
                        Element::TextSpace | Element::TextTab | Element::TextLineBreak
                            if current_paragraph.is_some() =>
                        {
                            if let Some(paragraph) = current_paragraph.as_mut() {
                                if !SELECT_ONE || slide_index == selected_index {
                                    Self::push_text_control(
                                        &*reader,
                                        element,
                                        element_type,
                                        paragraph,
                                    )?;
                                } else {
                                    let mut ignored = ParagraphText::default();
                                    Self::push_text_control(
                                        &*reader,
                                        element,
                                        element_type,
                                        &mut ignored,
                                    )?;
                                }
                            }
                        },
                        _ if in_notes => {},
                        Element::UnknownAnimation if in_slide => {
                            return Err(Error::InvalidFormat(format!(
                                "unknown ODF animation element '{}'",
                                String::from_utf8_lossy(element.local_name().as_ref()),
                            )));
                        },
                        Element::Animation(kind)
                            if in_slide
                                && shape_stack.is_empty()
                                && current_paragraph.is_none() =>
                        {
                            if !kind.allowed_at_page_root() {
                                return Err(Error::InvalidFormat(
                                    "anim:param is only valid below anim:command".to_string(),
                                ));
                            }
                            current_animations.push(Self::parse_animation_node(
                                reader,
                                collector,
                                element,
                                kind,
                                1,
                                &mut animation_node_count,
                            )?);
                            element_depth = Self::rewind_consumed_subtree(element_depth);
                        },
                        Element::SpreadsheetRoot if sheet_scope => {
                            spreadsheet_depth = Some(element_depth);
                        },
                        Element::Table
                            if sheet_scope
                                && shape_stack.is_empty()
                                && spreadsheet_depth
                                    .is_some_and(|depth| element_depth == depth + 1) =>
                        {
                            sheet_table_depth = Some(element_depth);
                            sheet_table_has_shapes = false;
                        },
                        Element::SheetShapes
                            if sheet_scope
                                && sheet_table_depth
                                    .is_some_and(|depth| element_depth == depth + 1) =>
                        {
                            if sheet_table_has_shapes {
                                return Err(Error::InvalidFormat(
                                    "table:table contains multiple table:shapes containers"
                                        .to_string(),
                                ));
                            }
                            sheet_table_has_shapes = true;
                            sheet_shapes_depth = Some(element_depth);
                            in_slide = true;
                        },
                        Element::Shape(shape_element) => {
                            let drawing_kind = Self::drawing_kind(shape_element);
                            shape_node_count =
                                shape_node_count.checked_add(1).ok_or_else(|| {
                                    Error::InvalidFormat("ODP shape count overflow".to_string())
                                })?;
                            if shape_node_count > 65_536 {
                                return Err(Error::InvalidFormat(
                                    "ODP document exceeds 65536 shapes".to_string(),
                                ));
                            }
                            if shape_stack.len() >= 64 {
                                return Err(Error::InvalidFormat(
                                    "ODP shape groups exceed 64 levels".to_string(),
                                ));
                            }
                            let hyperlink_applies = current_hyperlink.is_some()
                                && hyperlink_parent_depth == Some(shape_stack.len());
                            if hyperlink_applies && hyperlink_shape_seen {
                                return Err(Error::InvalidFormat(
                                    "draw:a must wrap exactly one drawing shape".to_string(),
                                ));
                            }
                            if in_slide && shape_stack.is_empty() {
                                if drawing_kind.is_three_dimensional()
                                    && drawing_kind != DrawingShapeKind::ThreeDimensionalScene
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D drawing objects require a dr3d:scene parent"
                                            .to_string(),
                                    ));
                                }
                                if current_hyperlink.is_some() && !hyperlink_applies {
                                    return Err(Error::InvalidFormat(
                                        "misplaced draw:a presentation hyperlink".to_string(),
                                    ));
                                }
                                let mut builder =
                                    Self::shape_builder(&*reader, element, shape_element)?;
                                if hyperlink_applies && let Some(hyperlink) = &current_hyperlink {
                                    builder.hyperlink = Some(hyperlink.clone());
                                    hyperlink_shape_seen = true;
                                }
                                shape_stack.push(builder);
                            } else if let Some(parent) = shape_stack.last() {
                                Self::validate_shape_parent(parent, drawing_kind)?;
                                if hyperlink_applies
                                    && parent.drawing_kind
                                        == Some(DrawingShapeKind::ThreeDimensionalScene)
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D scene children cannot be wrapped in draw:a".to_string(),
                                    ));
                                }
                                let mut builder =
                                    Self::shape_builder(&*reader, element, shape_element)?;
                                if hyperlink_applies && let Some(hyperlink) = &current_hyperlink {
                                    builder.hyperlink = Some(hyperlink.clone());
                                    hyperlink_shape_seen = true;
                                }
                                shape_stack.push(builder);
                            }
                        },
                        Element::Image if !shape_stack.is_empty() => {
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.shape_type = ShapeType::Picture;
                                builder.image_href =
                                    Self::get_attr(&*reader, element, XLINK_NAMESPACE, b"href")?;
                            }
                        },
                        Element::Table if !shape_stack.is_empty() => {
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.shape_type = ShapeType::Table;
                            }
                        },
                        Element::Object if !shape_stack.is_empty() => {
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.shape_type = ShapeType::GraphicFrame;
                            }
                        },
                        Element::Page
                        | Element::Notes
                        | Element::SheetShapes
                        | Element::SpreadsheetRoot
                        | Element::Image
                        | Element::Table
                        | Element::Object
                        | Element::Plugin
                        | Element::PluginParameter
                        | Element::DrawingHyperlink
                        | Element::EnhancedGeometry
                        | Element::EnhancedEquation
                        | Element::EnhancedHandle
                        | Element::EventListeners
                        | Element::EventListener
                        | Element::ScriptEventListener
                        | Element::Sound
                        | Element::TextParagraph
                        | Element::TextSpace
                        | Element::TextTab
                        | Element::TextLineBreak
                        | Element::Animation(_)
                        | Element::UnknownAnimation
                        | Element::LegacyAnimation(_)
                        | Element::Other => {},
                    }
                },
                Event::Text(text) if current_paragraph.is_some() => {
                    let decoded = Self::decode_text(text)?;
                    if (!SELECT_ONE || slide_index == selected_index)
                        && let Some(paragraph) = current_paragraph.as_mut()
                    {
                        paragraph.push_text(&decoded);
                    }
                },
                Event::Text(text) if in_media_plugin => {
                    let decoded = Self::decode_text(text)?;
                    if !decoded.trim().is_empty() {
                        return Err(Error::InvalidFormat(
                            "draw:plugin cannot contain text".to_string(),
                        ));
                    }
                },
                Event::Text(text)
                    if shape_stack.last().is_some_and(|builder| {
                        builder.drawing_kind.is_some_and(|kind| {
                            kind.is_three_dimensional()
                                && kind != DrawingShapeKind::ThreeDimensionalScene
                        })
                    }) =>
                {
                    let decoded = Self::decode_text(text)?;
                    if !decoded.trim().is_empty() {
                        return Err(Error::InvalidFormat(
                            "3D drawing elements cannot contain text".to_string(),
                        ));
                    }
                },
                Event::CData(text) if current_paragraph.is_some() => {
                    let decoded = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid presentation CDATA: {error}"))
                    })?;
                    if (!SELECT_ONE || slide_index == selected_index)
                        && let Some(paragraph) = current_paragraph.as_mut()
                    {
                        paragraph.push_text(&decoded);
                    }
                },
                Event::CData(text) if in_media_plugin => {
                    let decoded = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid media plugin CDATA: {error}"))
                    })?;
                    if !decoded.trim().is_empty() {
                        return Err(Error::InvalidFormat(
                            "draw:plugin cannot contain text".to_string(),
                        ));
                    }
                },
                Event::GeneralRef(reference) if current_paragraph.is_some() => {
                    let text = Self::decode_reference(reference)?;
                    if (!SELECT_ONE || slide_index == selected_index)
                        && let Some(paragraph) = current_paragraph.as_mut()
                    {
                        paragraph.push_text(&text);
                    }
                },
                Event::GeneralRef(_) if in_media_plugin => {
                    return Err(Error::InvalidFormat(
                        "draw:plugin cannot contain character references".to_string(),
                    ));
                },
                Event::GeneralRef(_)
                    if shape_stack.last().is_some_and(|builder| {
                        builder.drawing_kind.is_some_and(|kind| {
                            kind.is_three_dimensional()
                                && kind != DrawingShapeKind::ThreeDimensionalScene
                        })
                    }) =>
                {
                    return Err(Error::InvalidFormat(
                        "3D drawing elements cannot contain character references".to_string(),
                    ));
                },
                Event::CData(data)
                    if shape_stack.last().is_some_and(|builder| {
                        builder.drawing_kind.is_some_and(|kind| {
                            kind.is_three_dimensional()
                                && kind != DrawingShapeKind::ThreeDimensionalScene
                        })
                    }) && !data.iter().all(u8::is_ascii_whitespace) =>
                {
                    return Err(Error::InvalidFormat(
                        "3D drawing elements cannot contain CDATA text".to_string(),
                    ));
                },
                Event::Empty(element) => {
                    let element_type = Self::classify(ns_class, element.local_name().as_ref());
                    Self::validate_three_dimensional_child_element(
                        shape_stack.last(),
                        element_type,
                    )?;
                    if in_media_parameter {
                        return Err(Error::InvalidFormat(
                            "draw:param cannot contain child elements".to_string(),
                        ));
                    }
                    if in_media_plugin && !matches!(element_type, Element::PluginParameter) {
                        return Err(Error::InvalidFormat(
                            "draw:plugin can only contain draw:param elements".to_string(),
                        ));
                    }
                    match element_type {
                        Element::Page if !sheet_scope && !in_slide => {
                            let style_name =
                                Self::get_attr(&*reader, element, DRAW_NAMESPACE, b"style-name")?;
                            if !SELECT_ONE || slide_index == selected_index {
                                slides.push(Slide {
                                    title: None,
                                    text: String::new(),
                                    index: slide_index,
                                    notes: None,
                                    transition: None,
                                    animations: Vec::new(),
                                    legacy_animation: None,
                                    shapes: Vec::new(),
                                });
                                pending_transitions.push(style_name);
                            }
                            slide_index += 1;
                        },
                        Element::EnhancedGeometry if !shape_stack.is_empty() => {
                            if let Some(builder) = shape_stack.last_mut() {
                                if builder.drawing_kind != Some(DrawingShapeKind::CustomShape) {
                                    return Err(Error::InvalidFormat(
                                        "draw:enhanced-geometry requires draw:custom-shape"
                                            .to_string(),
                                    ));
                                }
                                if builder.enhanced_geometry.is_some() {
                                    return Err(Error::InvalidFormat(
                                        "draw:custom-shape contains multiple enhanced geometries"
                                            .to_string(),
                                    ));
                                }
                                builder.enhanced_geometry = Some(EnhancedGeometry {
                                    attributes: Self::exact_geometry_attributes(&*reader, element)?,
                                    children: Vec::new(),
                                });
                            }
                        },
                        Element::EnhancedGeometry
                        | Element::EnhancedEquation
                        | Element::EnhancedHandle
                            if in_slide =>
                        {
                            return Err(Error::InvalidFormat(
                                "misplaced custom-shape enhanced geometry".to_string(),
                            ));
                        },
                        Element::Plugin => {
                            if let Some(builder) = shape_stack.last_mut() {
                                if !builder.is_frame {
                                    return Err(Error::InvalidFormat(
                                        "draw:plugin must be contained directly by draw:frame"
                                            .to_string(),
                                    ));
                                }
                                if builder.media.is_some() {
                                    return Err(Error::InvalidFormat(
                                        "ODP frame contains multiple draw:plugin elements"
                                            .to_string(),
                                    ));
                                }
                                builder.shape_type = ShapeType::GraphicFrame;
                                builder.media = Some(Self::media_reference(&*reader, element)?);
                            } else if in_slide {
                                return Err(Error::InvalidFormat(
                                    "draw:plugin must be contained by a drawing shape".to_string(),
                                ));
                            }
                        },
                        Element::DrawingHyperlink if in_slide => {
                            return Err(Error::InvalidFormat(
                                "draw:a must wrap exactly one non-empty drawing shape".to_string(),
                            ));
                        },
                        Element::EventListeners if !shape_stack.is_empty() => {
                            if let Some(builder) = shape_stack.last_mut() {
                                if builder
                                    .drawing_kind
                                    .is_some_and(DrawingShapeKind::is_three_dimensional)
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D shapes cannot contain presentation event listeners"
                                            .to_string(),
                                    ));
                                }
                                if builder.event_listeners_seen {
                                    return Err(Error::InvalidFormat(
                                        "ODP shape contains multiple office:event-listeners elements"
                                            .to_string(),
                                    ));
                                }
                                builder.event_listeners_seen = true;
                            }
                        },
                        Element::EventListeners
                        | Element::EventListener
                        | Element::ScriptEventListener
                        | Element::Sound
                            if in_slide =>
                        {
                            return Err(Error::InvalidFormat(
                                "presentation event metadata must be contained by a shape's office:event-listeners"
                                    .to_string(),
                            ));
                        },
                        Element::PluginParameter
                            if in_media_plugin
                                && !in_media_parameter
                                && !shape_stack.is_empty() =>
                        {
                            if let Some(media) = shape_stack
                                .last_mut()
                                .and_then(|builder| builder.media.as_mut())
                            {
                                media.add_parameter(Self::media_parameter(&*reader, element)?)?;
                            }
                        },
                        _ if in_media_parameter => {
                            return Err(Error::InvalidFormat(
                                "draw:param cannot contain child elements".to_string(),
                            ));
                        },
                        _ if in_media_plugin => {
                            return Err(Error::InvalidFormat(
                                "draw:plugin can only contain draw:param elements".to_string(),
                            ));
                        },
                        Element::TextParagraph if in_slide => {
                            if !SELECT_ONE || slide_index == selected_index {
                                Self::push_parsed_paragraph(
                                    "",
                                    in_notes,
                                    &mut current_notes_text,
                                    &mut current_notes_has_paragraph,
                                    shape_stack.last_mut(),
                                    &mut current_slide_text,
                                    &mut current_slide_has_segment,
                                );
                            }
                        },
                        Element::TextSpace | Element::TextTab | Element::TextLineBreak
                            if current_paragraph.is_some() =>
                        {
                            if let Some(paragraph) = current_paragraph.as_mut() {
                                if !SELECT_ONE || slide_index == selected_index {
                                    Self::push_text_control(
                                        &*reader,
                                        element,
                                        element_type,
                                        paragraph,
                                    )?;
                                } else {
                                    let mut ignored = ParagraphText::default();
                                    Self::push_text_control(
                                        &*reader,
                                        element,
                                        element_type,
                                        &mut ignored,
                                    )?;
                                }
                            }
                        },
                        _ if in_notes => {},
                        Element::LegacyAnimation(kind) if in_slide => {
                            if kind != AnimationKind::Animations {
                                return Err(Error::InvalidFormat(
                                    "legacy presentation effects require a presentation:animations root"
                                        .to_string(),
                                ));
                            }
                            if current_legacy_animation.is_some() {
                                return Err(Error::InvalidFormat(
                                    "ODP slide contains multiple presentation:animations roots"
                                        .to_string(),
                                ));
                            }
                            legacy_animation_node_count =
                                legacy_animation_node_count.checked_add(1).ok_or_else(|| {
                                    Error::InvalidFormat(
                                        "legacy ODP animation node count overflow".to_string(),
                                    )
                                })?;
                            let root = AnimationNode::from_parsed(
                                kind,
                                Self::animation_attributes(&*reader, element)?,
                                Vec::new(),
                            );
                            validate_legacy_animation_root(&root)?;
                            current_legacy_animation = Some(root);
                        },
                        Element::UnknownAnimation if in_slide => {
                            return Err(Error::InvalidFormat(format!(
                                "unknown ODF animation element '{}'",
                                String::from_utf8_lossy(element.local_name().as_ref()),
                            )));
                        },
                        Element::Animation(kind)
                            if in_slide
                                && shape_stack.is_empty()
                                && current_paragraph.is_none() =>
                        {
                            if !kind.allowed_at_page_root() {
                                return Err(Error::InvalidFormat(
                                    "anim:param is only valid below anim:command".to_string(),
                                ));
                            }
                            animation_node_count =
                                animation_node_count.checked_add(1).ok_or_else(|| {
                                    Error::InvalidFormat(
                                        "ODP animation node count overflow".to_string(),
                                    )
                                })?;
                            if animation_node_count > 65_536 {
                                return Err(Error::InvalidFormat(
                                    "ODP animation tree exceeds 65536 nodes".to_string(),
                                ));
                            }
                            current_animations.push(Node::from_parsed(
                                kind,
                                Self::animation_attributes(&*reader, element)?,
                                Vec::new(),
                            ));
                        },
                        Element::Table
                            if sheet_scope
                                && shape_stack.is_empty()
                                && spreadsheet_depth
                                    .is_some_and(|depth| element_depth == depth) =>
                        {
                            slides.push(Slide {
                                title: None,
                                text: String::new(),
                                index: slide_index,
                                notes: None,
                                transition: None,
                                animations: Vec::new(),
                                legacy_animation: None,
                                shapes: Vec::new(),
                            });
                            slide_index += 1;
                        },
                        Element::SheetShapes
                            if sheet_scope
                                && sheet_table_depth
                                    .is_some_and(|depth| element_depth == depth) =>
                        {
                            if sheet_table_has_shapes {
                                return Err(Error::InvalidFormat(
                                    "table:table contains multiple table:shapes containers"
                                        .to_string(),
                                ));
                            }
                            sheet_table_has_shapes = true;
                        },
                        Element::Image => {
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.shape_type = ShapeType::Picture;
                                builder.image_href =
                                    Self::get_attr(&*reader, element, XLINK_NAMESPACE, b"href")?;
                            }
                        },
                        Element::Table => {
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.shape_type = ShapeType::Table;
                            }
                        },
                        Element::Object => {
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.shape_type = ShapeType::GraphicFrame;
                            }
                        },
                        Element::Shape(shape_element) if in_slide => {
                            let drawing_kind = Self::drawing_kind(shape_element);
                            shape_node_count =
                                shape_node_count.checked_add(1).ok_or_else(|| {
                                    Error::InvalidFormat("ODP shape count overflow".to_string())
                                })?;
                            if shape_node_count > 65_536 {
                                return Err(Error::InvalidFormat(
                                    "ODP document exceeds 65536 shapes".to_string(),
                                ));
                            }
                            let hyperlink_applies = current_hyperlink.is_some()
                                && hyperlink_parent_depth == Some(shape_stack.len());
                            if hyperlink_applies && hyperlink_shape_seen {
                                return Err(Error::InvalidFormat(
                                    "draw:a must wrap exactly one drawing shape".to_string(),
                                ));
                            }
                            let mut builder =
                                Self::shape_builder(&*reader, element, shape_element)?;
                            if hyperlink_applies && let Some(hyperlink) = &current_hyperlink {
                                builder.hyperlink = Some(hyperlink.clone());
                                hyperlink_shape_seen = true;
                            }
                            if let Some(parent) = shape_stack.last_mut() {
                                Self::validate_shape_parent(parent, drawing_kind)?;
                                if hyperlink_applies
                                    && parent.drawing_kind
                                        == Some(DrawingShapeKind::ThreeDimensionalScene)
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D scene children cannot be wrapped in draw:a".to_string(),
                                    ));
                                }
                                if !SELECT_ONE || slide_index == selected_index {
                                    parent.children.push(builder.build());
                                }
                            } else {
                                if drawing_kind.is_three_dimensional()
                                    && drawing_kind != DrawingShapeKind::ThreeDimensionalScene
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D drawing objects require a dr3d:scene parent"
                                            .to_string(),
                                    ));
                                }
                                if !SELECT_ONE || slide_index == selected_index {
                                    Self::finish_shape(
                                        builder,
                                        &mut current_slide_title,
                                        &mut current_slide_text,
                                        &mut current_slide_has_segment,
                                        &mut current_shapes,
                                        retain_text_shapes,
                                    );
                                }
                            }
                        },
                        Element::Page
                        | Element::Notes
                        | Element::SheetShapes
                        | Element::SpreadsheetRoot
                        | Element::Shape(_)
                        | Element::PluginParameter
                        | Element::DrawingHyperlink
                        | Element::EnhancedGeometry
                        | Element::EnhancedEquation
                        | Element::EnhancedHandle
                        | Element::EventListeners
                        | Element::EventListener
                        | Element::ScriptEventListener
                        | Element::Sound
                        | Element::TextParagraph
                        | Element::TextSpace
                        | Element::TextTab
                        | Element::TextLineBreak
                        | Element::Animation(_)
                        | Element::UnknownAnimation
                        | Element::LegacyAnimation(_)
                        | Element::Other => {},
                    }
                },
                Event::End(element) => {
                    element_depth = element_depth.saturating_sub(1);
                    let element_type = Self::classify(ns_class, element.local_name().as_ref());
                    if matches!(element_type, Element::TextParagraph)
                        && let Some(parsed_paragraph) = current_paragraph.take()
                    {
                        if !SELECT_ONE || slide_index == selected_index {
                            let paragraph = parsed_paragraph.finish();
                            Self::push_parsed_paragraph(
                                &paragraph,
                                in_notes,
                                &mut current_notes_text,
                                &mut current_notes_has_paragraph,
                                shape_stack.last_mut(),
                                &mut current_slide_text,
                                &mut current_slide_has_segment,
                            );
                        }
                        return Ok(());
                    }
                    if matches!(element_type, Element::Notes) {
                        in_notes = false;
                        return Ok(());
                    }
                    if matches!(element_type, Element::Plugin) {
                        in_media_plugin = false;
                        return Ok(());
                    }
                    if matches!(element_type, Element::PluginParameter) && in_media_parameter {
                        in_media_parameter = false;
                        return Ok(());
                    }
                    if in_notes {
                        return Ok(());
                    }
                    match element_type {
                        Element::DrawingHyperlink if current_hyperlink.is_some() => {
                            if hyperlink_parent_depth != Some(shape_stack.len())
                                || !hyperlink_shape_seen
                            {
                                return Err(Error::InvalidFormat(
                                    "draw:a must wrap exactly one complete drawing shape"
                                        .to_string(),
                                ));
                            }
                            current_hyperlink = None;
                            hyperlink_parent_depth = None;
                            hyperlink_shape_seen = false;
                        },
                        Element::Page if !sheet_scope => {
                            if in_slide {
                                if current_hyperlink.is_some() {
                                    return Err(Error::InvalidFormat(
                                        "unterminated draw:a presentation hyperlink".to_string(),
                                    ));
                                }
                                if !SELECT_ONE || slide_index == selected_index {
                                    slides.push(Slide {
                                        title: current_slide_title.take(),
                                        text: std::mem::take(&mut current_slide_text),
                                        index: slide_index,
                                        notes: (!current_notes_text.is_empty())
                                            .then(|| std::mem::take(&mut current_notes_text)),
                                        transition: None,
                                        animations: std::mem::take(&mut current_animations),
                                        legacy_animation: current_legacy_animation.take(),
                                        shapes: std::mem::take(&mut current_shapes),
                                    });
                                    pending_transitions.push(current_page_style.take());
                                } else {
                                    current_slide_title = None;
                                    current_slide_text.clear();
                                    current_notes_text.clear();
                                    current_page_style = None;
                                    current_animations.clear();
                                    current_legacy_animation = None;
                                    current_shapes.clear();
                                }
                                slide_index += 1;
                            }
                            current_slide_has_segment = false;
                            current_notes_has_paragraph = false;
                            in_slide = false;
                        },
                        Element::SpreadsheetRoot
                            if sheet_scope
                                && spreadsheet_depth
                                    .is_some_and(|depth| element_depth + 1 == depth) =>
                        {
                            spreadsheet_depth = None;
                        },
                        Element::SheetShapes
                            if sheet_scope
                                && sheet_shapes_depth
                                    .is_some_and(|depth| element_depth + 1 == depth) =>
                        {
                            if current_hyperlink.is_some() {
                                return Err(Error::InvalidFormat(
                                    "unterminated draw:a drawing hyperlink".to_string(),
                                ));
                            }
                            sheet_shapes_depth = None;
                            in_slide = false;
                        },
                        Element::Table
                            if sheet_scope
                                && sheet_table_depth
                                    .is_some_and(|depth| element_depth + 1 == depth) =>
                        {
                            slides.push(Slide {
                                title: None,
                                text: std::mem::take(&mut current_slide_text),
                                index: slide_index,
                                notes: None,
                                transition: None,
                                animations: Vec::new(),
                                legacy_animation: None,
                                shapes: std::mem::take(&mut current_shapes),
                            });
                            slide_index += 1;
                            sheet_table_depth = None;
                            current_slide_has_segment = false;
                        },
                        Element::Shape(_) => {
                            if let Some(builder) = shape_stack.pop() {
                                if let Some(parent) = shape_stack.last_mut() {
                                    if !SELECT_ONE || slide_index == selected_index {
                                        parent.children.push(builder.build());
                                    }
                                    return Ok(());
                                }
                                if !SELECT_ONE || slide_index == selected_index {
                                    Self::finish_shape(
                                        builder,
                                        &mut current_slide_title,
                                        &mut current_slide_text,
                                        &mut current_slide_has_segment,
                                        &mut current_shapes,
                                        retain_text_shapes,
                                    );
                                }
                            }
                        },
                        Element::Page
                        | Element::Notes
                        | Element::SheetShapes
                        | Element::SpreadsheetRoot
                        | Element::Image
                        | Element::Table
                        | Element::Object
                        | Element::Plugin
                        | Element::PluginParameter
                        | Element::DrawingHyperlink
                        | Element::EnhancedGeometry
                        | Element::EnhancedEquation
                        | Element::EnhancedHandle
                        | Element::EventListeners
                        | Element::EventListener
                        | Element::ScriptEventListener
                        | Element::Sound
                        | Element::TextParagraph
                        | Element::TextSpace
                        | Element::TextTab
                        | Element::TextLineBreak
                        | Element::Animation(_)
                        | Element::UnknownAnimation
                        | Element::LegacyAnimation(_)
                        | Element::Other => {},
                    }
                },
                Event::Eof => {},
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => {},
            }
            Ok(())
        };

        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buf)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            let ns_class = NsClass::from_resolve(&namespace);
            collector.feed(&reader, ns_class, &event);
            if let Some(error) = collector.take_error() {
                return Err(error);
            }
            if matches!(event, Event::Eof) {
                break;
            }
            if main_error.is_none() {
                let outcome = process(&mut reader, &mut collector, ns_class, &event);
                // A collection error recorded while nested subtree parsers ran
                // keeps transition-scan precedence over any slide-scan error
                // the same call surfaced.
                if let Some(error) = collector.take_error() {
                    return Err(error);
                }
                if let Err(error) = outcome {
                    if Self::is_xml_read_error(&error) {
                        return Err(error);
                    }
                    main_error = Some(error);
                }
            }
            buf.clear();
        }

        Self::merge_transition_style_definitions(&mut definitions, collector.finish());
        let (transition_styles, default_transition) = Self::resolve_transition_styles(definitions)?;
        if let Some(error) = main_error {
            return Err(error);
        }
        for (slide, style_name) in slides.iter_mut().zip(&pending_transitions) {
            let transition = style_name
                .as_deref()
                .and_then(|name| transition_styles.get(name))
                .unwrap_or(&default_transition)
                .clone();
            slide.transition = (!transition.is_empty()).then_some(transition);
        }
        Ok(slides)
    }
}

/// Incremental mirror of [`Parser::parse_transition_style_definitions`] for the
/// fused content.xml pass.
///
/// The fused pass feeds every tokenization event to this collector in stream
/// order — including events consumed by the nested subtree parsers — so the
/// collected definitions are exactly those the standalone pre-scan produced.
/// The first collection error is recorded rather than returned; the fused
/// driver surfaces it ahead of any recorded slide-scan error, matching the
/// historical pass order in which the transition scan ran to completion (or
/// failed) before slide parsing started.
const SINK_MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;
const SINK_MAX_TEXT_FRAGMENTS: usize = 1_000_000;
const SINK_MAX_TEXT_DEPTH: usize = 4096;
const SINK_MAX_SHAPES: usize = 65_536;
const SINK_MAX_SHAPE_DEPTH: usize = 64;
const SINK_MAX_SPACE_COUNT: usize = 1_000_000;
const SINK_MAX_ELEMENT_NAME_BYTES: usize = 1_048_576;
const SINK_MAX_OPEN_ELEMENT_NAME_BYTES: usize = 4 * 1_048_576;

/// Bounded accounting for decoded visible text on the fused presentation pass.
///
/// The ordinary slide projection keeps one complete model, whereas this path
/// keeps only the current slide. Charging before appending still prevents a
/// document with many simultaneously active nested shape strings from gaining
/// one full text budget per string.
struct SinkTextBudget {
    decoded_bytes: usize,
    fragments: usize,
}

impl SinkTextBudget {
    fn charge(&mut self, additional: usize) -> Result<()> {
        self.decoded_bytes = self
            .decoded_bytes
            .checked_add(additional)
            .ok_or_else(|| Error::InvalidFormat("ODP text size overflow".to_string()))?;
        if self.decoded_bytes > SINK_MAX_TEXT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "ODP text exceeds {SINK_MAX_TEXT_BYTES} bytes"
            )));
        }
        Ok(())
    }

    fn count_fragment(&mut self) -> Result<()> {
        self.fragments = self
            .fragments
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("ODP text fragment count overflow".to_string()))?;
        if self.fragments > SINK_MAX_TEXT_FRAGMENTS {
            return Err(Error::InvalidFormat(format!(
                "ODP text exceeds {SINK_MAX_TEXT_FRAGMENTS} paragraphs"
            )));
        }
        Ok(())
    }
}

#[derive(Default)]
struct SinkElementNameBudget {
    open_bytes: usize,
    open_lengths: Vec<usize>,
}

impl SinkElementNameBudget {
    fn validate_name(name: &[u8]) -> Result<usize> {
        if name.len() > SINK_MAX_ELEMENT_NAME_BYTES {
            return Err(Error::InvalidFormat(format!(
                "ODP element name exceeds {SINK_MAX_ELEMENT_NAME_BYTES} bytes"
            )));
        }
        Ok(name.len())
    }

    fn start(&mut self, name: &[u8]) -> Result<()> {
        let length = Self::validate_name(name)?;
        let total = self.open_bytes.checked_add(length).ok_or_else(|| {
            Error::InvalidFormat("ODP open element name size overflow".to_string())
        })?;
        if total > SINK_MAX_OPEN_ELEMENT_NAME_BYTES {
            return Err(Error::InvalidFormat(format!(
                "ODP open element names exceed {SINK_MAX_OPEN_ELEMENT_NAME_BYTES} bytes"
            )));
        }
        self.open_lengths
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODP sink open element names",
                source,
            })?;
        self.open_lengths.push(length);
        self.open_bytes = total;
        Ok(())
    }

    fn end(&mut self, name: &[u8]) -> Result<()> {
        Self::validate_name(name)?;
        let length = self.open_lengths.pop().ok_or_else(|| {
            Error::InvalidFormat("ODP sink element name stack underflow".to_string())
        })?;
        self.open_bytes = self.open_bytes.checked_sub(length).ok_or_else(|| {
            Error::InvalidFormat("ODP open element name size underflow".to_string())
        })?;
        Ok(())
    }

    fn is_empty(&self) -> bool {
        self.open_lengths.is_empty()
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum SinkShapeKind {
    TextBox,
    Placeholder,
    Other,
}

struct SinkShape {
    kind: SinkShapeKind,
    is_title: bool,
    paragraphs: SinkTextPart,
    children: Vec<SinkShape>,
}

impl SinkShape {
    fn new(kind: SinkShapeKind, is_title: bool) -> Self {
        Self {
            kind,
            is_title,
            paragraphs: SinkTextPart::new(),
            children: Vec::new(),
        }
    }

    fn push_paragraph(
        &mut self,
        text: String,
        budget: &mut SinkTextBudget,
        separator: &str,
    ) -> Result<()> {
        self.paragraphs.push(text, budget, separator)
    }
}

struct SinkTextPart {
    paragraphs: Vec<String>,
}

impl SinkTextPart {
    fn new() -> Self {
        Self {
            paragraphs: Vec::new(),
        }
    }

    fn push(&mut self, text: String, budget: &mut SinkTextBudget, separator: &str) -> Result<()> {
        if !self.paragraphs.is_empty() {
            budget.charge(separator.len())?;
        }
        budget.count_fragment()?;
        self.paragraphs
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODP sink text fragments",
                source,
            })?;
        self.paragraphs.push(text);
        Ok(())
    }

    fn append_part(
        &mut self,
        mut part: SinkTextPart,
        budget: &mut SinkTextBudget,
        separator: &str,
    ) -> Result<()> {
        if part.paragraphs.is_empty() {
            return Ok(());
        }
        if !self.paragraphs.is_empty() {
            budget.charge(separator.len())?;
        }
        self.paragraphs
            .try_reserve(part.paragraphs.len())
            .map_err(|source| Error::Allocation {
                resource: "ODP sink text fragments",
                source,
            })?;
        self.paragraphs.append(&mut part.paragraphs);
        Ok(())
    }

    fn trim_outer(mut self) -> Option<Self> {
        let mut first = 0;
        while first < self.paragraphs.len() && self.paragraphs[first].trim_start().is_empty() {
            first += 1;
        }
        let mut last = self.paragraphs.len();
        while last > first && self.paragraphs[last - 1].trim_end().is_empty() {
            last -= 1;
        }
        if first != 0 {
            self.paragraphs.drain(..first).for_each(drop);
        }
        self.paragraphs.truncate(last - first);
        if let Some(first) = self.paragraphs.first_mut() {
            let start = first.len().saturating_sub(first.trim_start().len());
            if start != 0 {
                first.drain(..start).for_each(drop);
            }
        }
        if let Some(last) = self.paragraphs.last_mut() {
            last.truncate(last.trim_end().len());
        }
        (!self.paragraphs.is_empty()).then_some(self)
    }
}

struct SinkSlideState {
    title: Option<SinkTextPart>,
    body: SinkTextPart,
    shapes: Vec<SinkShape>,
    shape_stack: Vec<SinkShape>,
}

impl SinkSlideState {
    fn new() -> Self {
        Self {
            title: None,
            body: SinkTextPart::new(),
            shapes: Vec::new(),
            shape_stack: Vec::new(),
        }
    }

    fn finish_shape(
        &mut self,
        shape: SinkShape,
        budget: &mut SinkTextBudget,
        separator: &str,
    ) -> Result<()> {
        if let Some(parent) = self.shape_stack.last_mut() {
            parent
                .children
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODP sink shape children",
                    source,
                })?;
            parent.children.push(shape);
            return Ok(());
        }

        if shape.is_title {
            self.title = Some(shape.paragraphs);
        } else if matches!(
            shape.kind,
            SinkShapeKind::TextBox | SinkShapeKind::Placeholder
        ) && shape
            .paragraphs
            .paragraphs
            .iter()
            .any(|text| !text.trim().is_empty())
        {
            self.body.append_part(shape.paragraphs, budget, separator)?;
        } else {
            self.shapes
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODP sink top-level shapes",
                    source,
                })?;
            self.shapes.push(shape);
        }
        Ok(())
    }

    fn append_body_paragraph(
        &mut self,
        text: String,
        budget: &mut SinkTextBudget,
        separator: &str,
    ) -> Result<()> {
        self.body.push(text, budget, separator)
    }

    fn into_parts(self, budget: &mut SinkTextBudget, separator: &str) -> Result<Vec<String>> {
        let Self {
            title,
            body,
            shapes,
            shape_stack: _,
        } = self;
        let mut output = Vec::new();
        output
            .try_reserve(shapes.len().saturating_add(2))
            .map_err(|source| Error::Allocation {
                resource: "ODP sink slide fragments",
                source,
            })?;
        if let Some(title) = title.and_then(SinkTextPart::trim_outer) {
            append_sink_part_fragments(&mut output, title, budget, separator)?;
        }
        if let Some(body) = body.trim_outer() {
            append_sink_part_fragments(&mut output, body, budget, separator)?;
        }
        for shape in shapes {
            collect_sink_shape_parts(shape, &mut output, budget, separator)?;
        }
        Ok(output)
    }
}

fn append_sink_part_fragments(
    output: &mut Vec<String>,
    part: SinkTextPart,
    budget: &mut SinkTextBudget,
    separator: &str,
) -> Result<()> {
    if !output.is_empty() {
        budget.charge(separator.len())?;
    }
    output
        .try_reserve(part.paragraphs.len())
        .map_err(|source| Error::Allocation {
            resource: "ODP sink slide fragments",
            source,
        })?;
    output.extend(part.paragraphs);
    Ok(())
}

fn collect_sink_shape_parts(
    shape: SinkShape,
    output: &mut Vec<String>,
    budget: &mut SinkTextBudget,
    separator: &str,
) -> Result<()> {
    let SinkShape {
        paragraphs,
        children,
        kind: _,
        is_title: _,
    } = shape;
    if let Some(paragraphs) = paragraphs.trim_outer() {
        append_sink_part_fragments(output, paragraphs, budget, separator)?;
    }
    for child in children {
        collect_sink_shape_parts(child, output, budget, separator)?;
    }
    Ok(())
}

fn sink_is_page(namespace: NsClass, local_name: &[u8]) -> bool {
    namespace == NsClass::Drawing && local_name == b"page"
}

fn sink_is_notes(namespace: NsClass, local_name: &[u8]) -> bool {
    namespace == NsClass::Presentation && local_name == b"notes"
}

fn sink_is_text_block(namespace: NsClass, local_name: &[u8]) -> bool {
    namespace == NsClass::Text && matches!(local_name, b"p" | b"h")
}

fn sink_shape_kind(namespace: NsClass, local_name: &[u8]) -> Option<SinkShapeKind> {
    if namespace == NsClass::Drawing {
        return match local_name {
            b"frame" => Some(SinkShapeKind::TextBox),
            b"rect" | b"ellipse" | b"line" | b"custom-shape" | b"circle" | b"path" | b"polygon"
            | b"polyline" | b"regular-polygon" | b"page-thumbnail" | b"measure" | b"caption"
            | b"connector" | b"control" | b"g" => Some(SinkShapeKind::Other),
            _ => None,
        };
    }
    if namespace == NsClass::Dr3d {
        return match local_name {
            b"scene" | b"light" | b"cube" | b"sphere" | b"extrude" | b"rotate" => {
                Some(SinkShapeKind::Other)
            },
            _ => None,
        };
    }
    None
}

fn sink_shape(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: NsClass,
) -> Result<SinkShape> {
    let presentation_class = sink_presentation_class(reader, element)?;
    let kind = if namespace == NsClass::Drawing && element.local_name().as_ref() == b"frame" {
        if !matches!(presentation_class, SinkPresentationClass::Absent) {
            SinkShapeKind::Placeholder
        } else {
            SinkShapeKind::TextBox
        }
    } else {
        SinkShapeKind::Other
    };
    Ok(SinkShape::new(
        kind,
        matches!(presentation_class, SinkPresentationClass::Title),
    ))
}

const SINK_MAX_ATTRIBUTES: usize = 256;
const SINK_MAX_ATTRIBUTE_VALUE_BYTES: usize = 1_048_576;

#[derive(Clone, Copy, PartialEq, Eq)]
enum SinkPresentationClass {
    Absent,
    Present,
    Title,
}

fn sink_presentation_class(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<SinkPresentationClass> {
    let mut attribute_count = 0usize;
    for result in element.attributes() {
        attribute_count = attribute_count.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("ODP shape attribute count overflow".to_string())
        })?;
        if attribute_count > SINK_MAX_ATTRIBUTES {
            return Err(Error::InvalidFormat(format!(
                "ODP shape exceeds {SINK_MAX_ATTRIBUTES} attributes"
            )));
        }
        let attribute = result.map_err(|error| {
            Error::InvalidFormat(format!("invalid ODP shape attribute: {error}"))
        })?;
        if attribute.value.len() > SINK_MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(Error::InvalidFormat(format!(
                "ODP shape attribute exceeds {SINK_MAX_ATTRIBUTE_VALUE_BYTES} bytes"
            )));
        }
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Bound(XmlNamespace(uri)) if uri == PRESENTATION_NAMESPACE)
            || local_name.as_ref() != b"class"
        {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid presentation:class value: {error}"))
            })?;
        if value.len() > SINK_MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(Error::InvalidFormat(format!(
                "presentation:class exceeds {SINK_MAX_ATTRIBUTE_VALUE_BYTES} bytes"
            )));
        }
        return Ok(if value.as_ref() == "title" {
            SinkPresentationClass::Title
        } else {
            SinkPresentationClass::Present
        });
    }
    Ok(SinkPresentationClass::Absent)
}

fn sink_text_space_count(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<usize> {
    let mut attribute_count = 0usize;
    let mut count = None;
    for result in element.attributes() {
        attribute_count = attribute_count.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("ODP text:s attribute count overflow".to_string())
        })?;
        if attribute_count > SINK_MAX_ATTRIBUTES {
            return Err(Error::InvalidFormat(format!(
                "ODP text:s exceeds {SINK_MAX_ATTRIBUTES} attributes"
            )));
        }
        let attribute = result.map_err(|error| {
            Error::InvalidFormat(format!("invalid ODP text:s attribute: {error}"))
        })?;
        if attribute.value.len() > SINK_MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(Error::InvalidFormat(format!(
                "ODP text:s attribute exceeds {SINK_MAX_ATTRIBUTE_VALUE_BYTES} bytes"
            )));
        }
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Bound(XmlNamespace(uri)) if uri == super::super::TEXT_NAMESPACE)
            || local_name.as_ref() != b"c"
        {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| Error::InvalidFormat(format!("invalid text:s count: {error}")))?;
        if value.len() > SINK_MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(Error::InvalidFormat(format!(
                "text:s count exceeds {SINK_MAX_ATTRIBUTE_VALUE_BYTES} bytes"
            )));
        }
        let parsed = value
            .parse::<usize>()
            .map_err(|_error| Error::InvalidFormat("invalid text:s count value".to_string()))?;
        if count.replace(parsed).is_some() {
            return Err(Error::InvalidFormat(
                "duplicate text:s count attribute".to_string(),
            ));
        }
    }
    Ok(count.unwrap_or(1))
}

fn normalized_xml10_decoded_len(raw: &[u8], context: &str) -> Result<usize> {
    std::str::from_utf8(raw).map_err(|error| {
        Error::InvalidFormat(format!("invalid presentation {context}: {error}"))
    })?;
    let mut length = raw.len();
    let mut index = 0;
    while index < raw.len() {
        if raw[index] == b'\r' {
            if raw.get(index + 1) == Some(&b'\n') {
                length = length
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("ODP text size overflow".to_string()))?;
                index += 2;
            } else {
                index += 1;
            }
        } else {
            index += 1;
        }
    }
    Ok(length)
}

fn sink_push_decoded_text(
    paragraph: &mut ParagraphText,
    decoded: &str,
    charged_len: usize,
) -> Result<()> {
    if decoded.len() != charged_len {
        return Err(Error::InvalidFormat(
            "ODP sink text precharge mismatch".to_string(),
        ));
    }
    paragraph
        .value
        .try_reserve(charged_len)
        .map_err(|source| Error::Allocation {
            resource: "ODP sink paragraph text",
            source,
        })?;
    paragraph.push_text(decoded);
    Ok(())
}

fn sink_push_xml_text(
    paragraph: &mut ParagraphText,
    raw: &[u8],
    context: &str,
    budget: &mut SinkTextBudget,
) -> Result<()> {
    let decoded_len = normalized_xml10_decoded_len(raw, context)?;
    budget.charge(decoded_len)?;
    let mut segment_start = 0usize;
    let mut index = 0usize;
    while index < raw.len() {
        if raw[index] == b'\r' {
            if segment_start != index {
                let segment = std::str::from_utf8(&raw[segment_start..index]).map_err(|error| {
                    Error::InvalidFormat(format!("invalid presentation {context}: {error}"))
                })?;
                sink_push_decoded_text(paragraph, segment, segment.len())?;
            }
            if raw.get(index + 1) == Some(&b'\n') {
                index += 1;
            }
            sink_push_decoded_text(paragraph, "\n", 1)?;
            index += 1;
            segment_start = index;
        } else {
            index += 1;
        }
    }
    if segment_start != raw.len() {
        let segment = std::str::from_utf8(&raw[segment_start..]).map_err(|error| {
            Error::InvalidFormat(format!("invalid presentation {context}: {error}"))
        })?;
        sink_push_decoded_text(paragraph, segment, segment.len())?;
    }
    Ok(())
}

fn sink_push_control(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local_name: &[u8],
    paragraph: &mut ParagraphText,
    budget: &mut SinkTextBudget,
) -> Result<()> {
    match local_name {
        b"s" => {
            let count = sink_text_space_count(reader, element)?;
            if count > SINK_MAX_SPACE_COUNT {
                return Err(Error::InvalidFormat(format!(
                    "text:s count exceeds {SINK_MAX_SPACE_COUNT}"
                )));
            }
            budget.charge(count)?;
            paragraph
                .value
                .try_reserve(count)
                .map_err(|source| Error::Allocation {
                    resource: "ODP sink paragraph spaces",
                    source,
                })?;
            paragraph.push_explicit(' ', count);
        },
        b"tab" => {
            budget.charge(1)?;
            paragraph
                .value
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODP sink paragraph tab",
                    source,
                })?;
            paragraph.push_explicit('\t', 1);
        },
        b"line-break" => {
            budget.charge(1)?;
            paragraph
                .value
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODP sink paragraph line break",
                    source,
                })?;
            paragraph.push_explicit('\n', 1);
        },
        _ => {},
    }
    Ok(())
}

fn finish_sink_slide<'options, 'output, W: Write + ?Sized>(
    slide: &mut Option<SinkSlideState>,
    writer: &mut SequentialTextWriter<'options, 'output, W>,
    budget: &mut SinkTextBudget,
    paragraph_separator: &str,
) -> std::result::Result<(), TextOutputError<Error>> {
    let Some(slide) = slide.take() else {
        return Ok(());
    };
    if !slide.shape_stack.is_empty() {
        return Err(writer.document_error(Error::InvalidFormat(
            "ODP sink shape stack is not empty at slide end".to_string(),
        )));
    }
    let parts = slide
        .into_parts(budget, paragraph_separator)
        .map_err(|error| writer.document_error(error))?;
    writer.write_joined_object::<Error, _, _>(
        TextObjectKind::Slide,
        || parts.iter().map(String::as_str),
        paragraph_separator,
    )
}

impl Parser {
    /// Feed one visible semantic text object per `draw:page` to a bounded sink.
    ///
    /// This is intentionally separate from the retained slide model path. It
    /// preserves the title/body/shape traversal used by [`Slide::all_text`]
    /// while retaining only the current slide and its nested shape text.
    pub(crate) fn write_slides_text_to<'options, 'output, W: Write + ?Sized>(
        xml_content: &str,
        writer: &mut SequentialTextWriter<'options, 'output, W>,
        paragraph_separator: &str,
    ) -> std::result::Result<(), TextOutputError<Error>> {
        let mut reader = NsReader::from_str(xml_content);
        let mut depth = 0usize;
        let mut slide_count = 0usize;
        let mut shape_count = 0usize;
        let mut notes_depth = 0usize;
        let mut in_slide = false;
        let mut slide: Option<SinkSlideState> = None;
        let mut paragraph: Option<ParagraphText> = None;
        let mut budget = SinkTextBudget {
            decoded_bytes: 0,
            fragments: 0,
        };
        let mut element_names = SinkElementNameBudget::default();

        loop {
            let (namespace, event) = reader.read_resolved_event().map_err(|error| {
                writer.document_error(Error::InvalidFormat(format!("XML parsing error: {error}")))
            })?;
            let ns_class = NsClass::from_resolve(&namespace);

            match event {
                Event::Start(element) => {
                    element_names
                        .start(element.name().as_ref())
                        .map_err(|error| writer.document_error(error))?;
                    depth = depth.checked_add(1).ok_or_else(|| {
                        writer.document_error(Error::InvalidFormat(
                            "ODP text nesting depth overflow".to_string(),
                        ))
                    })?;
                    if depth > SINK_MAX_TEXT_DEPTH {
                        return Err(writer.document_error(Error::InvalidFormat(format!(
                            "ODP text nesting exceeds {SINK_MAX_TEXT_DEPTH} levels"
                        ))));
                    }

                    if notes_depth > 0 {
                        notes_depth += 1;
                        continue;
                    }
                    if sink_is_notes(ns_class, element.local_name().as_ref()) {
                        notes_depth = 1;
                        continue;
                    }
                    if sink_is_page(ns_class, element.local_name().as_ref()) {
                        if in_slide {
                            if paragraph.is_some() {
                                return Err(writer.document_error(Error::InvalidFormat(
                                    "ODP text paragraph is open at slide boundary".to_string(),
                                )));
                            }
                            finish_sink_slide(
                                &mut slide,
                                writer,
                                &mut budget,
                                paragraph_separator,
                            )?;
                        }
                        slide_count = slide_count.checked_add(1).ok_or_else(|| {
                            writer.document_error(Error::InvalidFormat(
                                "ODP slide count overflow".to_string(),
                            ))
                        })?;
                        if slide_count > SINK_MAX_SHAPES {
                            return Err(writer.document_error(Error::InvalidFormat(
                                "ODP document exceeds 65536 slides".to_string(),
                            )));
                        }
                        slide = Some(SinkSlideState::new());
                        in_slide = true;
                        continue;
                    }
                    if !in_slide {
                        continue;
                    }
                    if sink_shape_kind(ns_class, element.local_name().as_ref()).is_some() {
                        shape_count = shape_count.checked_add(1).ok_or_else(|| {
                            writer.document_error(Error::InvalidFormat(
                                "ODP shape count overflow".to_string(),
                            ))
                        })?;
                        if shape_count > SINK_MAX_SHAPES {
                            return Err(writer.document_error(Error::InvalidFormat(
                                "ODP document exceeds 65536 shapes".to_string(),
                            )));
                        }
                        let state = slide.as_mut().ok_or_else(|| {
                            writer.document_error(Error::InvalidFormat(
                                "ODP sink slide state is missing".to_string(),
                            ))
                        })?;
                        if state.shape_stack.len() >= SINK_MAX_SHAPE_DEPTH {
                            return Err(writer.document_error(Error::InvalidFormat(
                                "ODP shape groups exceed 64 levels".to_string(),
                            )));
                        }
                        let shape = sink_shape(&reader, &element, ns_class)
                            .map_err(|error| writer.document_error(error))?;
                        state.shape_stack.try_reserve(1).map_err(|source| {
                            writer.document_error(Error::Allocation {
                                resource: "ODP sink shape stack",
                                source,
                            })
                        })?;
                        state.shape_stack.push(shape);
                        continue;
                    }
                    if sink_is_text_block(ns_class, element.local_name().as_ref()) {
                        if paragraph.is_some() {
                            return Err(writer.document_error(Error::InvalidFormat(
                                "nested ODP text paragraphs are not supported".to_string(),
                            )));
                        }
                        paragraph = Some(ParagraphText::default());
                    } else if ns_class == NsClass::Text
                        && paragraph.is_some()
                        && matches!(element.local_name().as_ref(), b"s" | b"tab" | b"line-break")
                    {
                        let paragraph = paragraph.as_mut().ok_or_else(|| {
                            writer.document_error(Error::InvalidFormat(
                                "ODP sink paragraph state is missing".to_string(),
                            ))
                        })?;
                        sink_push_control(
                            &reader,
                            &element,
                            element.local_name().as_ref(),
                            paragraph,
                            &mut budget,
                        )
                        .map_err(|error| writer.document_error(error))?;
                    }
                },
                Event::Empty(element) => {
                    SinkElementNameBudget::validate_name(element.name().as_ref())
                        .map_err(|error| writer.document_error(error))?;
                    if notes_depth > 0 {
                        continue;
                    }
                    if sink_is_page(ns_class, element.local_name().as_ref()) {
                        if in_slide {
                            if paragraph.is_some() {
                                return Err(writer.document_error(Error::InvalidFormat(
                                    "ODP text paragraph is open at slide boundary".to_string(),
                                )));
                            }
                            finish_sink_slide(
                                &mut slide,
                                writer,
                                &mut budget,
                                paragraph_separator,
                            )?;
                        }
                        slide_count = slide_count.checked_add(1).ok_or_else(|| {
                            writer.document_error(Error::InvalidFormat(
                                "ODP slide count overflow".to_string(),
                            ))
                        })?;
                        if slide_count > SINK_MAX_SHAPES {
                            return Err(writer.document_error(Error::InvalidFormat(
                                "ODP document exceeds 65536 slides".to_string(),
                            )));
                        }
                        slide = Some(SinkSlideState::new());
                        finish_sink_slide(&mut slide, writer, &mut budget, paragraph_separator)?;
                        in_slide = false;
                        continue;
                    }
                    if !in_slide {
                        continue;
                    }
                    if sink_shape_kind(ns_class, element.local_name().as_ref()).is_some() {
                        shape_count = shape_count.checked_add(1).ok_or_else(|| {
                            writer.document_error(Error::InvalidFormat(
                                "ODP shape count overflow".to_string(),
                            ))
                        })?;
                        if shape_count > SINK_MAX_SHAPES {
                            return Err(writer.document_error(Error::InvalidFormat(
                                "ODP document exceeds 65536 shapes".to_string(),
                            )));
                        }
                        let state = slide.as_mut().ok_or_else(|| {
                            writer.document_error(Error::InvalidFormat(
                                "ODP sink slide state is missing".to_string(),
                            ))
                        })?;
                        let shape = sink_shape(&reader, &element, ns_class)
                            .map_err(|error| writer.document_error(error))?;
                        state
                            .finish_shape(shape, &mut budget, paragraph_separator)
                            .map_err(|error| writer.document_error(error))?;
                        continue;
                    }
                    if sink_is_text_block(ns_class, element.local_name().as_ref()) {
                        if paragraph.is_some() {
                            return Err(writer.document_error(Error::InvalidFormat(
                                "nested ODP text paragraphs are not supported".to_string(),
                            )));
                        }
                        paragraph = Some(ParagraphText::default());
                        let value = paragraph
                            .take()
                            .ok_or_else(|| {
                                writer.document_error(Error::InvalidFormat(
                                    "ODP sink paragraph state is missing".to_string(),
                                ))
                            })?
                            .finish();
                        let state = slide.as_mut().ok_or_else(|| {
                            writer.document_error(Error::InvalidFormat(
                                "ODP sink slide state is missing".to_string(),
                            ))
                        })?;
                        if let Some(shape) = state.shape_stack.last_mut() {
                            shape
                                .push_paragraph(value, &mut budget, paragraph_separator)
                                .map_err(|error| writer.document_error(error))?;
                        } else {
                            state
                                .append_body_paragraph(value, &mut budget, paragraph_separator)
                                .map_err(|error| writer.document_error(error))?;
                        }
                        continue;
                    }
                    if ns_class == NsClass::Text
                        && paragraph.is_some()
                        && matches!(element.local_name().as_ref(), b"s" | b"tab" | b"line-break")
                    {
                        let paragraph = paragraph.as_mut().ok_or_else(|| {
                            writer.document_error(Error::InvalidFormat(
                                "ODP sink paragraph state is missing".to_string(),
                            ))
                        })?;
                        sink_push_control(
                            &reader,
                            &element,
                            element.local_name().as_ref(),
                            paragraph,
                            &mut budget,
                        )
                        .map_err(|error| writer.document_error(error))?;
                    }
                },
                Event::End(element) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        writer.document_error(Error::InvalidFormat(
                            "ODP text element stack underflow".to_string(),
                        ))
                    })?;
                    element_names
                        .end(element.name().as_ref())
                        .map_err(|error| writer.document_error(error))?;
                    if notes_depth > 0 {
                        notes_depth -= 1;
                        continue;
                    }
                    if sink_is_text_block(ns_class, element.local_name().as_ref()) {
                        let Some(paragraph_value) = paragraph.take() else {
                            continue;
                        };
                        let value = paragraph_value.finish();
                        let state = slide.as_mut().ok_or_else(|| {
                            writer.document_error(Error::InvalidFormat(
                                "ODP sink slide state is missing".to_string(),
                            ))
                        })?;
                        if let Some(shape) = state.shape_stack.last_mut() {
                            shape
                                .push_paragraph(value, &mut budget, paragraph_separator)
                                .map_err(|error| writer.document_error(error))?;
                        } else {
                            state
                                .append_body_paragraph(value, &mut budget, paragraph_separator)
                                .map_err(|error| writer.document_error(error))?;
                        }
                    } else if sink_shape_kind(ns_class, element.local_name().as_ref()).is_some() {
                        let state = slide.as_mut().ok_or_else(|| {
                            writer.document_error(Error::InvalidFormat(
                                "ODP sink slide state is missing".to_string(),
                            ))
                        })?;
                        if let Some(shape) = state.shape_stack.pop() {
                            state
                                .finish_shape(shape, &mut budget, paragraph_separator)
                                .map_err(|error| writer.document_error(error))?;
                        }
                    } else if sink_is_page(ns_class, element.local_name().as_ref()) {
                        if paragraph.is_some() {
                            return Err(writer.document_error(Error::InvalidFormat(
                                "ODP text paragraph is open at slide end".to_string(),
                            )));
                        }
                        finish_sink_slide(&mut slide, writer, &mut budget, paragraph_separator)?;
                        in_slide = false;
                    }
                },
                Event::Text(text) if paragraph.is_some() => {
                    let paragraph = paragraph.as_mut().ok_or_else(|| {
                        writer.document_error(Error::InvalidFormat(
                            "ODP sink paragraph state is missing".to_string(),
                        ))
                    })?;
                    sink_push_xml_text(paragraph, text.as_ref(), "text content", &mut budget)
                        .map_err(|error| writer.document_error(error))?;
                },
                Event::CData(text) if paragraph.is_some() => {
                    let paragraph = paragraph.as_mut().ok_or_else(|| {
                        writer.document_error(Error::InvalidFormat(
                            "ODP sink paragraph state is missing".to_string(),
                        ))
                    })?;
                    sink_push_xml_text(paragraph, text.as_ref(), "CDATA", &mut budget)
                        .map_err(|error| writer.document_error(error))?;
                },
                Event::GeneralRef(reference) if paragraph.is_some() => {
                    let paragraph = paragraph.as_mut().ok_or_else(|| {
                        writer.document_error(Error::InvalidFormat(
                            "ODP sink paragraph state is missing".to_string(),
                        ))
                    })?;
                    sink_push_reference(&reference, paragraph, &mut budget)
                        .map_err(|error| writer.document_error(error))?;
                },
                Event::Eof => break,
                _ => {},
            }
        }

        if depth != 0
            || !element_names.is_empty()
            || notes_depth != 0
            || paragraph.is_some()
            || in_slide
            || slide.is_some()
        {
            return Err(writer.document_error(Error::InvalidFormat(
                "incomplete ODP presentation text XML structure".to_string(),
            )));
        }
        Ok(())
    }
}

fn sink_push_reference(
    reference: &BytesRef<'_>,
    paragraph: &mut ParagraphText,
    budget: &mut SinkTextBudget,
) -> Result<()> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid presentation character reference: {error}"))
    })? {
        let mut encoded = [0_u8; 4];
        let value = character.encode_utf8(&mut encoded);
        budget.charge(value.len())?;
        return sink_push_decoded_text(paragraph, value, value.len());
    }

    let value = match reference.as_ref() {
        b"amp" => "&",
        b"lt" => "<",
        b"gt" => ">",
        b"quot" => "\"",
        b"apos" => "'",
        _ => {
            return Err(Error::InvalidFormat(
                "unsupported presentation entity reference".to_string(),
            ));
        },
    };
    budget.charge(value.len())?;
    sink_push_decoded_text(paragraph, value, value.len())
}

#[derive(Default)]
pub(super) struct TransitionStyleCollector {
    result: TransitionStyles,
    current: Option<(Option<String>, bool, TransitionStyleDefinition)>,
    in_properties: bool,
    error: Option<Error>,
}

impl TransitionStyleCollector {
    /// Feeds one event; a poisoned collector ignores all further events, just
    /// like the standalone scan never reads past its first error.
    pub(super) fn feed(&mut self, reader: &NsReader<&[u8]>, namespace: NsClass, event: &Event<'_>) {
        if self.error.is_none()
            && let Err(error) = self.apply(reader, namespace, event)
        {
            self.error = Some(error);
        }
    }

    /// Takes the recorded collection error, if any.
    pub(super) fn take_error(&mut self) -> Option<Error> {
        self.error.take()
    }

    /// Consumes the collector into the collected definitions.
    pub(super) fn finish(self) -> TransitionStyles {
        self.result
    }

    /// The match arms of [`Parser::parse_transition_style_definitions`],
    /// transcribed one-to-one with the same arm order and guards.
    fn apply(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: NsClass,
        event: &Event<'_>,
    ) -> Result<()> {
        match event {
            Event::Start(element)
                if namespace == NsClass::Style
                    && matches!(element.local_name().as_ref(), b"style" | b"default-style") =>
            {
                let mut attributes = ElementAttrs::new(element);
                let family = attributes.get(reader, STYLE_NAMESPACE, b"family")?;
                let is_drawing_page = family.as_deref() == Some("drawing-page");
                let name = attributes.get(reader, STYLE_NAMESPACE, b"name")?;
                let parent = attributes.get(reader, STYLE_NAMESPACE, b"parent-style-name")?;
                self.current = Some((
                    name,
                    is_drawing_page,
                    TransitionStyleDefinition {
                        parent,
                        transition: Transition::new(),
                    },
                ));
            },
            Event::Empty(element)
                if namespace == NsClass::Style
                    && matches!(element.local_name().as_ref(), b"style" | b"default-style") =>
            {
                let mut attributes = ElementAttrs::new(element);
                let family = attributes.get(reader, STYLE_NAMESPACE, b"family")?;
                if family.as_deref() == Some("drawing-page") {
                    let name = attributes.get(reader, STYLE_NAMESPACE, b"name")?;
                    let definition = TransitionStyleDefinition {
                        parent: attributes.get(reader, STYLE_NAMESPACE, b"parent-style-name")?,
                        transition: Transition::new(),
                    };
                    if let Some(style_name) = name {
                        self.result.named.insert(style_name, definition);
                    } else {
                        self.result.default = definition.transition;
                    }
                }
            },
            Event::Start(element) | Event::Empty(element)
                if self.current.as_ref().is_some_and(|(_, family, _)| *family)
                    && namespace == NsClass::Style
                    && element.local_name().as_ref() == b"drawing-page-properties" =>
            {
                if let Some((_, _, definition)) = self.current.as_mut() {
                    Parser::parse_transition_properties(
                        reader,
                        element,
                        &mut definition.transition,
                    )?;
                }
                self.in_properties = matches!(event, Event::Start(_));
            },
            Event::Start(element) | Event::Empty(element)
                if self.in_properties
                    && namespace == NsClass::Presentation
                    && element.local_name().as_ref() == b"sound" =>
            {
                if let Some((_, _, definition)) = self.current.as_mut() {
                    definition.transition.sound =
                        Some(Parser::parse_transition_sound(reader, element)?);
                }
            },
            Event::End(element)
                if namespace == NsClass::Style
                    && element.local_name().as_ref() == b"drawing-page-properties" =>
            {
                self.in_properties = false;
            },
            Event::End(element)
                if namespace == NsClass::Style
                    && matches!(element.local_name().as_ref(), b"style" | b"default-style") =>
            {
                if let Some((name, is_drawing_page, definition)) = self.current.take()
                    && is_drawing_page
                {
                    if let Some(style_name) = name {
                        self.result.named.insert(style_name, definition);
                    } else {
                        self.result.default = definition.transition;
                    }
                }
                self.in_properties = false;
            },
            _ => {},
        }
        Ok(())
    }
}
