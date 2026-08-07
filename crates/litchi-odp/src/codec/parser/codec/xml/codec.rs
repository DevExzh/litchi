//! XML token and event traversal for the ODP parser.

use super::super::*;

impl Parser {
    pub(super) fn parse_animation_node(
        reader: &mut NsReader<&[u8]>,
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
            match event {
                Event::Start(ref child) | Event::Empty(ref child) => {
                    if !Self::is_namespace(&namespace, ANIMATION_NAMESPACE_BYTES) {
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
                            child,
                            child_kind,
                            depth + 1,
                            node_count,
                        )?
                    };
                    children.push(node);
                },
                Event::End(ref end) => {
                    if !Self::is_namespace(&namespace, ANIMATION_NAMESPACE_BYTES)
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
                    let text = Self::decode_text(text)?;
                    if !text.trim().is_empty() {
                        return Err(Error::InvalidFormat(format!(
                            "anim:{} cannot contain text",
                            kind.local_name()
                        )));
                    }
                },
                Event::CData(ref text) => {
                    let text = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid animation CDATA: {error}"))
                    })?;
                    if !text.trim().is_empty() {
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
                _ => {},
            }
            buffer.clear();
        }
    }

    pub(super) fn parse_legacy_animation_node(
        reader: &mut NsReader<&[u8]>,
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
            match event {
                Event::Start(ref child) | Event::Empty(ref child) => {
                    if !Self::is_namespace(&namespace, PRESENTATION_NAMESPACE) {
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
                            child,
                            child_kind,
                            depth + 1,
                            node_count,
                        )?
                    };
                    children.push(node);
                },
                Event::End(ref end)
                    if Self::is_namespace(&namespace, PRESENTATION_NAMESPACE)
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
                _ => {},
            }
            buffer.clear();
        }
    }

    pub(super) fn parse_enhanced_geometry(
        reader: &mut NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<EnhancedGeometry> {
        let attributes = Self::exact_geometry_attributes(reader, element)?;
        let mut children = Vec::new();
        let mut handle_seen = false;
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) =
                reader
                    .read_resolved_event_into(&mut buffer)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid enhanced geometry XML: {error}"))
                    })?;
            match event {
                Event::Start(ref child) | Event::Empty(ref child)
                    if Self::is_namespace(&namespace, DRAW_NAMESPACE)
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
                            DRAW_NAMESPACE,
                            child.local_name().as_ref(),
                            kind.element_name(),
                        )?;
                    }
                },
                Event::End(ref end)
                    if Self::is_namespace(&namespace, DRAW_NAMESPACE)
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
                _ => {
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
        let event_name = Self::required_attr(
            reader,
            element,
            SCRIPT_NAMESPACE,
            b"event-name",
            "script:event-name",
        )?;
        let language = Self::required_attr(
            reader,
            element,
            SCRIPT_NAMESPACE,
            b"language",
            "script:language",
        )?;
        let macro_name = Self::get_attr(reader, element, SCRIPT_NAMESPACE, b"macro-name")?;
        let href = Self::get_attr(reader, element, XLINK_NAMESPACE, b"href")?;
        let link_type = Self::get_attr(reader, element, XLINK_NAMESPACE, b"type")?;
        if href.is_some() {
            Self::require_simple_xlink(reader, element, "script:event-listener")?;
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
                Self::get_attr(reader, element, XLINK_NAMESPACE, b"actuate")?,
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
        let event_name = Self::required_attr(
            reader,
            element,
            SCRIPT_NAMESPACE,
            b"event-name",
            "script:event-name",
        )?;
        let action = Action::parse(&Self::required_attr(
            reader,
            element,
            PRESENTATION_NAMESPACE,
            b"action",
            "presentation:action",
        )?)?;
        let mut listener = EventListener::new(event_name, action)?;
        listener.effect = Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"effect")?
            .map(Effect::new)
            .transpose()?;
        listener.direction = Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"direction")?
            .map(EffectDirection::new)
            .transpose()?;
        listener.speed = Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"speed")?
            .map(|value| Speed::parse(&value))
            .transpose()?;
        listener.start_scale =
            Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"start-scale")?;
        listener.href = Self::get_attr(reader, element, XLINK_NAMESPACE, b"href")?;
        let link_type = Self::get_attr(reader, element, XLINK_NAMESPACE, b"type")?;
        if listener.href.is_some() {
            Self::require_simple_xlink(reader, element, "presentation:event-listener")?;
        } else if link_type.is_some() {
            return Err(Error::InvalidFormat(
                "presentation:event-listener xlink:type requires xlink:href".to_string(),
            ));
        }
        listener.show_embed =
            match Self::get_attr(reader, element, XLINK_NAMESPACE, b"show")?.as_deref() {
                None => false,
                Some("embed") => true,
                Some(value) => {
                    return Err(Error::InvalidFormat(format!(
                        "invalid presentation:event-listener xlink:show '{value}'"
                    )));
                },
            };
        listener.actuate_on_request = Self::parse_on_request(
            Self::get_attr(reader, element, XLINK_NAMESPACE, b"actuate")?,
            "presentation:event-listener",
        )?;
        listener.verb = Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"verb")?
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
        namespace_uri: &[u8],
        local_name: &[u8],
        description: &str,
    ) -> Result<()> {
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            match event {
                Event::End(ref end)
                    if Self::is_namespace(&namespace, namespace_uri)
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
                _ => {
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
        mut listener: EventListener,
    ) -> Result<EventListener> {
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            match event {
                Event::Start(ref element) | Event::Empty(ref element)
                    if Self::is_namespace(&namespace, PRESENTATION_NAMESPACE)
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
                            PRESENTATION_NAMESPACE,
                            b"sound",
                            "presentation:sound",
                        )?;
                    }
                },
                Event::End(ref end)
                    if Self::is_namespace(&namespace, PRESENTATION_NAMESPACE)
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
                _ => {
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
    ) -> Result<Vec<ShapeEventListener>> {
        let mut listeners = Vec::new();
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            match event {
                Event::Start(ref element) | Event::Empty(ref element)
                    if Self::is_namespace(&namespace, SCRIPT_NAMESPACE)
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
                            SCRIPT_NAMESPACE,
                            b"event-listener",
                            "script:event-listener",
                        )?;
                    }
                    listeners.push(ShapeEventListener::Script(listener));
                },
                Event::Start(ref element) | Event::Empty(ref element)
                    if Self::is_namespace(&namespace, PRESENTATION_NAMESPACE)
                        && element.local_name().as_ref() == b"event-listener" =>
                {
                    if listeners.len() >= 4096 {
                        return Err(Error::InvalidFormat(
                            "ODP shape exceeds 4096 event listeners".to_string(),
                        ));
                    }
                    let listener = Self::presentation_event_listener(reader, element)?;
                    let listener = if matches!(event, Event::Start(_)) {
                        Self::parse_listener_body(reader, listener)?
                    } else {
                        listener
                    };
                    listeners.push(ShapeEventListener::Action(Box::new(listener)));
                },
                Event::End(ref end)
                    if Self::is_namespace(&namespace, OFFICE_NAMESPACE)
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
                _ => {
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
                    let family = Self::get_attr(&reader, element, STYLE_NAMESPACE, b"family")?;
                    let is_drawing_page = family.as_deref() == Some("drawing-page");
                    let name = Self::get_attr(&reader, element, STYLE_NAMESPACE, b"name")?;
                    let parent =
                        Self::get_attr(&reader, element, STYLE_NAMESPACE, b"parent-style-name")?;
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
                    let family = Self::get_attr(&reader, element, STYLE_NAMESPACE, b"family")?;
                    if family.as_deref() == Some("drawing-page") {
                        let name = Self::get_attr(&reader, element, STYLE_NAMESPACE, b"name")?;
                        let definition = TransitionStyleDefinition {
                            parent: Self::get_attr(
                                &reader,
                                element,
                                STYLE_NAMESPACE,
                                b"parent-style-name",
                            )?,
                            transition: Transition::new(),
                        };
                        if let Some(name) = name {
                            result.named.insert(name, definition);
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
                    let (_, _, definition) = current.as_mut().expect("style checked above");
                    Self::parse_transition_properties(
                        &reader,
                        element,
                        &mut definition.transition,
                    )?;
                    in_properties = matches!(event, Event::Start(_));
                },
                Event::Start(ref element) | Event::Empty(ref element)
                    if in_properties
                        && Self::is_namespace(&namespace, PRESENTATION_NAMESPACE)
                        && element.local_name().as_ref() == b"sound" =>
                {
                    let (_, _, definition) = current.as_mut().expect("properties require style");
                    definition.transition.sound =
                        Some(Self::parse_transition_sound(&reader, element)?);
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
                        if let Some(name) = name {
                            result.named.insert(name, definition);
                        } else {
                            result.default = definition.transition;
                        }
                    }
                    in_properties = false;
                },
                Event::Eof => break,
                _ => {},
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
        Self::parse_pages_with_styles(
            xml_content,
            styles_xml,
            false,
            ShapeContainerScope::DrawPages,
        )
    }

    /// Parse drawing pages while retaining title and text-box frames as shapes.
    #[allow(dead_code, reason = "reserved for the dedicated ODG facade")]
    pub(crate) fn parse_drawing_pages(
        xml_content: &str,
        styles_xml: Option<&str>,
    ) -> Result<Vec<Slide>> {
        Self::parse_pages_with_styles(
            xml_content,
            styles_xml,
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
        let tables = Self::parse_pages_with_styles(
            xml_content,
            None,
            true,
            ShapeContainerScope::SpreadsheetTables,
        )?;
        Ok(tables.into_iter().map(|table| table.shapes).collect())
    }

    pub(super) fn parse_pages_with_styles(
        xml_content: &str,
        styles_xml: Option<&str>,
        retain_text_shapes: bool,
        container_scope: ShapeContainerScope,
    ) -> Result<Vec<Slide>> {
        let sheet_scope = container_scope == ShapeContainerScope::SpreadsheetTables;
        let (transition_styles, default_transition) =
            Self::resolved_transition_styles(xml_content, styles_xml)?;
        let mut reader = NsReader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut slides = Vec::new();

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
        let mut current_transition: Option<Transition> = None;
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

        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buf)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            match event {
                Event::Start(ref element) => {
                    element_depth = element_depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("XML element depth overflow".to_string())
                    })?;
                    let element_type = Self::classify(&namespace, element.local_name().as_ref());
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
                                slides.push(Slide {
                                    title: current_slide_title.take(),
                                    text: std::mem::take(&mut current_slide_text),
                                    index: slide_index,
                                    notes: (!current_notes_text.is_empty())
                                        .then(|| std::mem::take(&mut current_notes_text)),
                                    transition: current_transition.take(),
                                    animations: std::mem::take(&mut current_animations),
                                    legacy_animation: current_legacy_animation.take(),
                                    shapes: std::mem::take(&mut current_shapes),
                                });
                                slide_index += 1;
                            }
                            current_slide_title = None;
                            current_slide_has_segment = false;
                            current_notes_has_paragraph = false;
                            let style_name =
                                Self::get_attr(&reader, element, DRAW_NAMESPACE, b"style-name")?;
                            let transition = style_name
                                .as_deref()
                                .and_then(|name| transition_styles.get(name))
                                .unwrap_or(&default_transition)
                                .clone();
                            current_transition = (!transition.is_empty()).then_some(transition);
                            in_slide = true;
                        },
                        Element::Notes if in_slide => in_notes = true,
                        Element::EnhancedGeometry if !shape_stack.is_empty() => {
                            let builder = shape_stack.last().expect("shape checked above");
                            if builder.drawing_kind != Some(DrawingShapeKind::CustomShape) {
                                return Err(Error::InvalidFormat(
                                    "draw:enhanced-geometry requires draw:custom-shape".to_string(),
                                ));
                            }
                            if builder.enhanced_geometry.is_some() {
                                return Err(Error::InvalidFormat(
                                    "draw:custom-shape contains multiple enhanced geometries"
                                        .to_string(),
                                ));
                            }
                            let geometry = Self::parse_enhanced_geometry(&mut reader, element)?;
                            element_depth = Self::rewind_consumed_subtree(element_depth);
                            shape_stack
                                .last_mut()
                                .expect("shape checked above")
                                .enhanced_geometry = Some(geometry);
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
                                &mut reader,
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
                            let builder = shape_stack.last_mut().expect("shape checked above");
                            if !builder.is_frame {
                                return Err(Error::InvalidFormat(
                                    "draw:plugin must be contained directly by draw:frame"
                                        .to_string(),
                                ));
                            }
                            if builder.media.is_some() {
                                return Err(Error::InvalidFormat(
                                    "ODP frame contains multiple draw:plugin elements".to_string(),
                                ));
                            }
                            builder.shape_type = ShapeType::GraphicFrame;
                            builder.media = Some(Self::media_reference(&reader, element)?);
                            in_media_plugin = true;
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
                            shape_stack
                                .last_mut()
                                .and_then(|builder| builder.media.as_mut())
                                .expect("media plugin state checked above")
                                .add_parameter(Self::media_parameter(&reader, element)?)?;
                            in_media_parameter = true;
                        },
                        Element::DrawingHyperlink
                            if in_slide && !in_notes && current_hyperlink.is_none() =>
                        {
                            current_hyperlink = Some(Self::drawing_hyperlink(&reader, element)?);
                            hyperlink_parent_depth = Some(shape_stack.len());
                            hyperlink_shape_seen = false;
                        },
                        Element::DrawingHyperlink if in_slide => {
                            return Err(Error::InvalidFormat(
                                "nested or misplaced draw:a presentation hyperlink".to_string(),
                            ));
                        },
                        Element::EventListeners if !shape_stack.is_empty() => {
                            let builder = shape_stack.last_mut().expect("shape checked above");
                            if builder
                                .drawing_kind
                                .is_some_and(|kind| kind.is_three_dimensional())
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
                            builder.event_listeners = Self::parse_event_listeners(&mut reader)?;
                            element_depth = Self::rewind_consumed_subtree(element_depth);
                            builder.event_listeners_seen = true;
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
                            Self::push_text_control(
                                &reader,
                                element,
                                element_type,
                                current_paragraph.as_mut().expect("paragraph checked above"),
                            )?;
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
                                &mut reader,
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
                                    Self::shape_builder(&reader, element, shape_element)?;
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
                                    Self::shape_builder(&reader, element, shape_element)?;
                                if hyperlink_applies && let Some(hyperlink) = &current_hyperlink {
                                    builder.hyperlink = Some(hyperlink.clone());
                                    hyperlink_shape_seen = true;
                                }
                                shape_stack.push(builder);
                            }
                        },
                        Element::Image if !shape_stack.is_empty() => {
                            let builder = shape_stack.last_mut().expect("shape checked above");
                            builder.shape_type = ShapeType::Picture;
                            builder.image_href =
                                Self::get_attr(&reader, element, XLINK_NAMESPACE, b"href")?;
                        },
                        Element::Table if !shape_stack.is_empty() => {
                            shape_stack
                                .last_mut()
                                .expect("shape checked above")
                                .shape_type = ShapeType::Table;
                        },
                        Element::Object if !shape_stack.is_empty() => {
                            shape_stack
                                .last_mut()
                                .expect("shape checked above")
                                .shape_type = ShapeType::GraphicFrame;
                        },
                        _ => {},
                    }
                },
                Event::Text(ref text) if current_paragraph.is_some() => {
                    let text = Self::decode_text(text)?;
                    current_paragraph
                        .as_mut()
                        .expect("paragraph checked above")
                        .push_text(&text);
                },
                Event::Text(ref text) if in_media_plugin => {
                    let text = Self::decode_text(text)?;
                    if !text.trim().is_empty() {
                        return Err(Error::InvalidFormat(
                            "draw:plugin cannot contain text".to_string(),
                        ));
                    }
                },
                Event::Text(ref text)
                    if shape_stack.last().is_some_and(|builder| {
                        builder.drawing_kind.is_some_and(|kind| {
                            kind.is_three_dimensional()
                                && kind != DrawingShapeKind::ThreeDimensionalScene
                        })
                    }) =>
                {
                    let text = Self::decode_text(text)?;
                    if !text.trim().is_empty() {
                        return Err(Error::InvalidFormat(
                            "3D drawing elements cannot contain text".to_string(),
                        ));
                    }
                },
                Event::CData(ref text) if current_paragraph.is_some() => {
                    let decoded = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid presentation CDATA: {error}"))
                    })?;
                    current_paragraph
                        .as_mut()
                        .expect("paragraph checked above")
                        .push_text(&decoded);
                },
                Event::CData(ref text) if in_media_plugin => {
                    let decoded = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid media plugin CDATA: {error}"))
                    })?;
                    if !decoded.trim().is_empty() {
                        return Err(Error::InvalidFormat(
                            "draw:plugin cannot contain text".to_string(),
                        ));
                    }
                },
                Event::GeneralRef(ref reference) if current_paragraph.is_some() => {
                    let text = Self::decode_reference(reference)?;
                    current_paragraph
                        .as_mut()
                        .expect("paragraph checked above")
                        .push_text(&text);
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
                Event::CData(ref data)
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
                Event::Empty(ref element) => {
                    let element_type = Self::classify(&namespace, element.local_name().as_ref());
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
                                Self::get_attr(&reader, element, DRAW_NAMESPACE, b"style-name")?;
                            let transition = style_name
                                .as_deref()
                                .and_then(|name| transition_styles.get(name))
                                .unwrap_or(&default_transition)
                                .clone();
                            slides.push(Slide {
                                title: None,
                                text: String::new(),
                                index: slide_index,
                                notes: None,
                                transition: (!transition.is_empty()).then_some(transition),
                                animations: Vec::new(),
                                legacy_animation: None,
                                shapes: Vec::new(),
                            });
                            slide_index += 1;
                        },
                        Element::EnhancedGeometry if !shape_stack.is_empty() => {
                            let builder = shape_stack.last_mut().expect("shape checked above");
                            if builder.drawing_kind != Some(DrawingShapeKind::CustomShape) {
                                return Err(Error::InvalidFormat(
                                    "draw:enhanced-geometry requires draw:custom-shape".to_string(),
                                ));
                            }
                            if builder.enhanced_geometry.is_some() {
                                return Err(Error::InvalidFormat(
                                    "draw:custom-shape contains multiple enhanced geometries"
                                        .to_string(),
                                ));
                            }
                            builder.enhanced_geometry = Some(EnhancedGeometry {
                                attributes: Self::exact_geometry_attributes(&reader, element)?,
                                children: Vec::new(),
                            });
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
                                builder.media = Some(Self::media_reference(&reader, element)?);
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
                            let builder = shape_stack.last_mut().expect("shape checked above");
                            if builder
                                .drawing_kind
                                .is_some_and(|kind| kind.is_three_dimensional())
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
                            shape_stack
                                .last_mut()
                                .and_then(|builder| builder.media.as_mut())
                                .expect("media plugin state checked above")
                                .add_parameter(Self::media_parameter(&reader, element)?)?;
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
                            Self::push_parsed_paragraph(
                                "",
                                in_notes,
                                &mut current_notes_text,
                                &mut current_notes_has_paragraph,
                                shape_stack.last_mut(),
                                &mut current_slide_text,
                                &mut current_slide_has_segment,
                            );
                        },
                        Element::TextSpace | Element::TextTab | Element::TextLineBreak
                            if current_paragraph.is_some() =>
                        {
                            Self::push_text_control(
                                &reader,
                                element,
                                element_type,
                                current_paragraph.as_mut().expect("paragraph checked above"),
                            )?;
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
                                Self::animation_attributes(&reader, element)?,
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
                                Self::animation_attributes(&reader, element)?,
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
                                    Self::get_attr(&reader, element, XLINK_NAMESPACE, b"href")?;
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
                            let mut builder = Self::shape_builder(&reader, element, shape_element)?;
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
                                parent.children.push(builder.build());
                            } else {
                                if drawing_kind.is_three_dimensional()
                                    && drawing_kind != DrawingShapeKind::ThreeDimensionalScene
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D drawing objects require a dr3d:scene parent"
                                            .to_string(),
                                    ));
                                }
                                Self::finish_shape(
                                    builder,
                                    &mut current_slide_title,
                                    &mut current_slide_text,
                                    &mut current_slide_has_segment,
                                    &mut current_shapes,
                                    retain_text_shapes,
                                );
                            }
                        },
                        _ => {},
                    }
                },
                Event::End(ref element) => {
                    element_depth = element_depth.saturating_sub(1);
                    let element_type = Self::classify(&namespace, element.local_name().as_ref());
                    if matches!(element_type, Element::TextParagraph) && current_paragraph.is_some()
                    {
                        let paragraph = current_paragraph
                            .take()
                            .expect("paragraph checked above")
                            .finish();
                        Self::push_parsed_paragraph(
                            &paragraph,
                            in_notes,
                            &mut current_notes_text,
                            &mut current_notes_has_paragraph,
                            shape_stack.last_mut(),
                            &mut current_slide_text,
                            &mut current_slide_has_segment,
                        );
                        buf.clear();
                        continue;
                    }
                    if matches!(element_type, Element::Notes) {
                        in_notes = false;
                        buf.clear();
                        continue;
                    }
                    if matches!(element_type, Element::Plugin) {
                        in_media_plugin = false;
                        buf.clear();
                        continue;
                    }
                    if matches!(element_type, Element::PluginParameter) && in_media_parameter {
                        in_media_parameter = false;
                        buf.clear();
                        continue;
                    }
                    if in_notes {
                        buf.clear();
                        continue;
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
                                slides.push(Slide {
                                    title: current_slide_title.take(),
                                    text: std::mem::take(&mut current_slide_text),
                                    index: slide_index,
                                    notes: (!current_notes_text.is_empty())
                                        .then(|| std::mem::take(&mut current_notes_text)),
                                    transition: current_transition.take(),
                                    animations: std::mem::take(&mut current_animations),
                                    legacy_animation: current_legacy_animation.take(),
                                    shapes: std::mem::take(&mut current_shapes),
                                });
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
                                    parent.children.push(builder.build());
                                    buf.clear();
                                    continue;
                                }
                                Self::finish_shape(
                                    builder,
                                    &mut current_slide_title,
                                    &mut current_slide_text,
                                    &mut current_slide_has_segment,
                                    &mut current_shapes,
                                    retain_text_shapes,
                                );
                            }
                        },
                        _ => {},
                    }
                },
                Event::Eof => break,
                _ => {},
            }
            buf.clear();
        }

        Ok(slides)
    }
}
