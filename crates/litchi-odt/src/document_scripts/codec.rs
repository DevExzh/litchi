//! Bounded XML codec for inert ODF document script declarations.

use litchi_core::{Error, Result};
use litchi_odf_common::{ResolvedReader, XmlStreamLimits};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::ResolveResult,
};

use super::model::{
    EmbeddedScript, EventListener, NamespaceDeclaration, ScriptBinding, ScriptEventListener,
    Scripts,
};
use super::{
    MAX_DOCUMENT_XML_BYTES, MAX_LISTENER_COUNT, MAX_SCRIPT_COUNT, MAX_TEXT_BYTES, MAX_VALUE_BYTES,
    MAX_XML_DEPTH, OFFICE_NAMESPACE, PRESENTATION_NAMESPACE, SCRIPT_NAMESPACE, XLINK_NAMESPACE,
    invalid,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Script,
    Presentation,
    Xlink,
    Other,
}

fn namespace_text(namespace: &[u8]) -> Result<&str> {
    std::str::from_utf8(namespace).map_err(|error| {
        Error::InvalidFormat(format!("invalid static namespace encoding: {error}"))
    })
}

pub(super) fn write_scripts(scripts: &Scripts, limits: &XmlStreamLimits) -> Result<String> {
    scripts.validate_for_output(limits)?;
    let capacity = scripts_size(scripts)?;
    ensure_serialized_size(capacity, limits)?;
    let mut output = String::with_capacity(capacity);
    append_scripts_open(&mut output, &scripts.namespace_declarations)?;
    output.push('>');

    for script in &scripts.scripts {
        output.push_str("<office:script script:language=\"");
        escape_attribute(&mut output, &script.language);
        if script.content_xml.is_empty() {
            output.push_str("\"/>");
        } else {
            output.push_str("\">");
            output.push_str(&script.content_xml);
            output.push_str("</office:script>");
        }
    }

    if !scripts.event_listeners.is_empty() {
        output.push_str("<office:event-listeners>");
        for listener in &scripts.event_listeners {
            match listener {
                EventListener::Script(listener) => {
                    output.push_str("<script:event-listener script:event-name=\"");
                    escape_attribute(&mut output, &listener.event_name);
                    output.push_str("\" script:language=\"");
                    escape_attribute(&mut output, &listener.language);
                    match &listener.binding {
                        ScriptBinding::MacroName(value) => {
                            output.push_str("\" script:macro-name=\"");
                            escape_attribute(&mut output, value);
                            output.push_str("\"/>");
                        },
                        ScriptBinding::Linked { href } => {
                            output.push_str("\" xlink:type=\"simple\" xlink:href=\"");
                            escape_attribute(&mut output, href);
                            output.push_str("\" xlink:actuate=\"onRequest\"/>");
                        },
                    }
                },
                EventListener::PresentationXml(xml) => output.push_str(xml),
            }
        }
        output.push_str("</office:event-listeners>");
    }
    output.push_str("</office:scripts>");
    if u64::try_from(output.len()).unwrap_or(u64::MAX) > limits.max_bytes() {
        return invalid("serialized office:scripts exceeds the configured XML limit");
    }
    Ok(output)
}

fn append_scripts_open(output: &mut String, declarations: &[NamespaceDeclaration]) -> Result<()> {
    output.push_str("<office:scripts");
    namespace_attribute(output, Some("office"), namespace_text(OFFICE_NAMESPACE)?);
    namespace_attribute(output, Some("script"), namespace_text(SCRIPT_NAMESPACE)?);
    namespace_attribute(
        output,
        Some("presentation"),
        namespace_text(PRESENTATION_NAMESPACE)?,
    );
    namespace_attribute(output, Some("xlink"), namespace_text(XLINK_NAMESPACE)?);
    for declaration in declarations {
        if matches!(
            declaration.prefix.as_deref(),
            Some("office" | "script" | "presentation" | "xlink")
        ) {
            continue;
        }
        namespace_attribute(output, declaration.prefix.as_deref(), &declaration.uri);
    }
    Ok(())
}

pub(super) fn emitted_namespace_count(declarations: &[NamespaceDeclaration]) -> usize {
    4 + declarations
        .iter()
        .filter(|declaration| {
            !matches!(
                declaration.prefix.as_deref(),
                Some("office" | "script" | "presentation" | "xlink")
            )
        })
        .count()
}

fn scripts_size(scripts: &Scripts) -> Result<usize> {
    let mut size = b"<office:scripts".len();
    add_size(
        &mut size,
        namespace_attribute_size(Some("office"), namespace_text(OFFICE_NAMESPACE)?)?,
    )?;
    add_size(
        &mut size,
        namespace_attribute_size(Some("script"), namespace_text(SCRIPT_NAMESPACE)?)?,
    )?;
    add_size(
        &mut size,
        namespace_attribute_size(
            Some("presentation"),
            namespace_text(PRESENTATION_NAMESPACE)?,
        )?,
    )?;
    add_size(
        &mut size,
        namespace_attribute_size(Some("xlink"), namespace_text(XLINK_NAMESPACE)?)?,
    )?;
    for declaration in &scripts.namespace_declarations {
        if matches!(
            declaration.prefix.as_deref(),
            Some("office" | "script" | "presentation" | "xlink")
        ) {
            continue;
        }
        add_size(
            &mut size,
            namespace_attribute_size(declaration.prefix.as_deref(), &declaration.uri)?,
        )?;
    }
    add_size(&mut size, 1)?;

    for script in &scripts.scripts {
        add_size(&mut size, b"<office:script script:language=\"".len())?;
        add_size(&mut size, escaped_attribute_len(&script.language)?)?;
        if script.content_xml.is_empty() {
            add_size(&mut size, b"\"/>".len())?;
        } else {
            add_size(&mut size, b"\">".len())?;
            add_size(&mut size, script.content_xml.len())?;
            add_size(&mut size, b"</office:script>".len())?;
        }
    }

    if !scripts.event_listeners.is_empty() {
        add_size(&mut size, b"<office:event-listeners>".len())?;
        for listener in &scripts.event_listeners {
            match listener {
                EventListener::Script(listener) => {
                    add_size(
                        &mut size,
                        b"<script:event-listener script:event-name=\"".len(),
                    )?;
                    add_size(&mut size, escaped_attribute_len(&listener.event_name)?)?;
                    add_size(&mut size, b"\" script:language=\"".len())?;
                    add_size(&mut size, escaped_attribute_len(&listener.language)?)?;
                    match &listener.binding {
                        ScriptBinding::MacroName(value) => {
                            add_size(&mut size, b"\" script:macro-name=\"".len())?;
                            add_size(&mut size, escaped_attribute_len(value)?)?;
                            add_size(&mut size, b"\"/>".len())?;
                        },
                        ScriptBinding::Linked { href } => {
                            add_size(&mut size, b"\" xlink:type=\"simple\" xlink:href=\"".len())?;
                            add_size(&mut size, escaped_attribute_len(href)?)?;
                            add_size(&mut size, b"\" xlink:actuate=\"onRequest\"/>".len())?;
                        },
                    }
                },
                EventListener::PresentationXml(xml) => add_size(&mut size, xml.len())?,
            }
        }
        add_size(&mut size, b"</office:event-listeners>".len())?;
    }
    add_size(&mut size, b"</office:scripts>".len())?;
    Ok(size)
}

fn namespace_attribute_size(prefix: Option<&str>, uri: &str) -> Result<usize> {
    let mut size = b" xmlns".len();
    if let Some(prefix) = prefix {
        add_size(&mut size, 1)?;
        add_size(&mut size, prefix.len())?;
    }
    add_size(&mut size, 2)?;
    add_size(&mut size, escaped_attribute_len(uri)?)?;
    add_size(&mut size, 1)?;
    Ok(size)
}

fn escaped_attribute_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |size, character| {
        let encoded = match character {
            '&' => 5,
            '<' => 4,
            '"' => 6,
            '\r' | '\n' => 5,
            '\t' => 4,
            _ => character.len_utf8(),
        };
        size.checked_add(encoded)
            .ok_or_else(|| Error::InvalidFormat("script XML output size overflow".to_string()))
    })
}

fn add_size(size: &mut usize, amount: usize) -> Result<()> {
    *size = size
        .checked_add(amount)
        .ok_or_else(|| Error::InvalidFormat("script XML output size overflow".to_string()))?;
    Ok(())
}

fn ensure_serialized_size(size: usize, limits: &XmlStreamLimits) -> Result<()> {
    if u64::try_from(size).unwrap_or(u64::MAX) > limits.max_bytes() {
        return invalid("serialized office:scripts exceeds the configured XML limit");
    }
    Ok(())
}

pub fn parse_scripts(xml: &str, limits: &XmlStreamLimits) -> Result<Option<Scripts>> {
    if xml.len() > MAX_DOCUMENT_XML_BYTES
        || u64::try_from(xml.len()).unwrap_or(u64::MAX) > limits.max_bytes()
    {
        return invalid(format!(
            "ODF XML exceeds the {MAX_DOCUMENT_XML_BYTES} byte script-inventory limit"
        ));
    }
    let mut reader = ResolvedReader::from_xml_document(xml, *limits).map_err(|error| {
        litchi_odf_common::map_resolved_reader_error(error, "invalid ODF script XML")
    })?;
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut root_namespaces = Vec::new();
    let mut result = None;

    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader.read_resolved_event().map_err(|error| {
            litchi_odf_common::map_resolved_reader_error(error, "invalid ODF script XML")
        })?;
        let namespace = namespace_kind(&namespace);
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    root_namespaces = namespace_declarations(&reader, &element)?;
                }
                if bound_to(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"scripts"
                {
                    if depth != 1 {
                        return invalid(
                            "office:scripts must be a direct child of the document root",
                        );
                    }
                    if result.is_some() {
                        return invalid("ODF document contains multiple office:scripts elements");
                    }
                    let mut in_scope_namespaces = root_namespaces.clone();
                    merge_namespace_declarations(
                        &mut in_scope_namespaces,
                        namespace_declarations(&reader, &element)?,
                    );
                    result = Some(parse_scripts_element(
                        &mut reader,
                        xml,
                        &in_scope_namespaces,
                        event_end,
                        limits,
                    )?);
                } else {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("ODF XML depth overflow".to_string())
                    })?;
                    if depth > MAX_XML_DEPTH.min(limits.max_depth()) {
                        return invalid(format!("ODF XML exceeds the {MAX_XML_DEPTH} depth limit"));
                    }
                }
            },
            Event::Empty(element) => {
                if depth == 0 {
                    root_namespaces = namespace_declarations(&reader, &element)?;
                }
                if bound_to(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"scripts"
                {
                    if depth != 1 {
                        return invalid(
                            "office:scripts must be a direct child of the document root",
                        );
                    }
                    if result.is_some() {
                        return invalid("ODF document contains multiple office:scripts elements");
                    }
                    let mut in_scope_namespaces = root_namespaces.clone();
                    merge_namespace_declarations(
                        &mut in_scope_namespaces,
                        namespace_declarations(&reader, &element)?,
                    );
                    result = Some(Scripts {
                        namespace_declarations: in_scope_namespaces,
                        ..Scripts::default()
                    });
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid ODF XML depth".to_string()))?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODF script metadata"),
            Event::Eof => break,
            _ => {},
        }
        if event_start > event_end || event_end > xml.len() {
            return invalid("invalid ODF XML reader position");
        }
    }
    Ok(result)
}

fn parse_scripts_element(
    reader: &mut ResolvedReader<'_>,
    xml: &str,
    root_namespaces: &[NamespaceDeclaration],
    _content_start: usize,
    limits: &XmlStreamLimits,
) -> Result<Scripts> {
    let mut scripts = Vec::new();
    let mut event_listeners = Vec::new();
    let mut listener_container_seen = false;
    let mut text_bytes = 0usize;
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader.read_resolved_event().map_err(|error| {
            litchi_odf_common::map_resolved_reader_error(error, "invalid office:scripts XML")
        })?;
        let namespace = namespace_kind(&namespace);
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(element)
                if bound_to(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"script" =>
            {
                if listener_container_seen {
                    return invalid("office:script must precede office:event-listeners");
                }
                if scripts.len() >= MAX_SCRIPT_COUNT {
                    return invalid(format!(
                        "office:scripts exceeds the {MAX_SCRIPT_COUNT} script limit"
                    ));
                }
                let language = required_script_language(reader, &element)?;
                let content_xml = capture_inner_xml(reader, xml, event_end, b"script", limits)?;
                text_bytes = checked_text_bytes(text_bytes, language.len())?;
                text_bytes = checked_text_bytes(text_bytes, content_xml.len())?;
                scripts.push(EmbeddedScript {
                    language,
                    content_xml,
                });
            },
            Event::Empty(element)
                if bound_to(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"script" =>
            {
                if listener_container_seen {
                    return invalid("office:script must precede office:event-listeners");
                }
                if scripts.len() >= MAX_SCRIPT_COUNT {
                    return invalid(format!(
                        "office:scripts exceeds the {MAX_SCRIPT_COUNT} script limit"
                    ));
                }
                let language = required_script_language(reader, &element)?;
                text_bytes = checked_text_bytes(text_bytes, language.len())?;
                scripts.push(EmbeddedScript {
                    language,
                    content_xml: String::new(),
                });
            },
            Event::Start(element)
                if bound_to(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"event-listeners" =>
            {
                if listener_container_seen {
                    return invalid("office:scripts contains duplicate office:event-listeners");
                }
                listener_container_seen = true;
                event_listeners = parse_event_listeners(reader, xml, &mut text_bytes, limits)?;
            },
            Event::Empty(element)
                if bound_to(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"event-listeners" =>
            {
                if listener_container_seen {
                    return invalid("office:scripts contains duplicate office:event-listeners");
                }
                listener_container_seen = true;
            },
            Event::End(element)
                if bound_to(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"scripts" =>
            {
                break;
            },
            Event::Text(text) => {
                let value = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid script text: {error}"))
                })?;
                if !value.trim().is_empty() {
                    return invalid("office:scripts cannot contain text");
                }
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in office:scripts"),
            Event::Eof => return invalid("unterminated office:scripts element"),
            _ => return invalid("unsupported child in office:scripts"),
        }
        if event_start > event_end || event_end > xml.len() {
            return invalid("invalid office:scripts reader position");
        }
    }

    Ok(Scripts {
        scripts,
        event_listeners,
        namespace_declarations: root_namespaces.to_vec(),
    })
}

fn parse_event_listeners(
    reader: &mut ResolvedReader<'_>,
    xml: &str,
    text_bytes: &mut usize,
    limits: &XmlStreamLimits,
) -> Result<Vec<EventListener>> {
    let mut listeners = Vec::new();
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader.read_resolved_event().map_err(|error| {
            litchi_odf_common::map_resolved_reader_error(error, "invalid event-listener XML")
        })?;
        let namespace = namespace_kind(&namespace);
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Empty(element)
                if bound_to(&namespace, SCRIPT_NAMESPACE)
                    && element.local_name().as_ref() == b"event-listener" =>
            {
                ensure_listener_capacity(listeners.len())?;
                let listener = parse_script_listener(reader, &element)?;
                *text_bytes = checked_text_bytes(*text_bytes, script_listener_bytes(&listener))?;
                listeners.push(EventListener::Script(listener));
            },
            Event::Start(element)
                if bound_to(&namespace, SCRIPT_NAMESPACE)
                    && element.local_name().as_ref() == b"event-listener" =>
            {
                ensure_listener_capacity(listeners.len())?;
                let listener = parse_script_listener(reader, &element)?;
                require_empty_element(reader, b"event-listener", SCRIPT_NAMESPACE)?;
                *text_bytes = checked_text_bytes(*text_bytes, script_listener_bytes(&listener))?;
                listeners.push(EventListener::Script(listener));
            },
            Event::Empty(element)
                if bound_to(&namespace, PRESENTATION_NAMESPACE)
                    && element.local_name().as_ref() == b"event-listener" =>
            {
                ensure_listener_capacity(listeners.len())?;
                let raw = xml
                    .get(event_start..event_end)
                    .ok_or_else(|| {
                        Error::InvalidFormat("invalid presentation listener range".to_string())
                    })?
                    .to_string();
                *text_bytes = checked_text_bytes(*text_bytes, raw.len())?;
                listeners.push(EventListener::PresentationXml(raw));
            },
            Event::Start(element)
                if bound_to(&namespace, PRESENTATION_NAMESPACE)
                    && element.local_name().as_ref() == b"event-listener" =>
            {
                ensure_listener_capacity(listeners.len())?;
                let raw = capture_full_xml(reader, xml, event_start, b"event-listener", limits)?;
                *text_bytes = checked_text_bytes(*text_bytes, raw.len())?;
                listeners.push(EventListener::PresentationXml(raw));
            },
            Event::End(element)
                if bound_to(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"event-listeners" =>
            {
                break;
            },
            Event::Text(text) => {
                let value = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid listener text: {error}"))
                })?;
                if !value.trim().is_empty() {
                    return invalid("office:event-listeners cannot contain text");
                }
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in event listeners"),
            Event::Eof => return invalid("unterminated office:event-listeners element"),
            _ => return invalid("unsupported child in office:event-listeners"),
        }
    }
    Ok(listeners)
}

fn parse_script_listener(
    reader: &ResolvedReader<'_>,
    element: &BytesStart<'_>,
) -> Result<ScriptEventListener> {
    let mut event_name = None;
    let mut language = None;
    let mut macro_name = None;
    let mut xlink_type = None;
    let mut href = None;
    let mut actuate = None;

    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid listener attribute: {error}"))
        })?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| Error::InvalidFormat(format!("invalid listener attribute: {error}")))?
            .into_owned();
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&namespace);
        let slot = if bound_to(&namespace, SCRIPT_NAMESPACE) {
            match local.as_ref() {
                b"event-name" => &mut event_name,
                b"language" => &mut language,
                b"macro-name" => &mut macro_name,
                _ => return invalid("unsupported script:event-listener attribute"),
            }
        } else if bound_to(&namespace, XLINK_NAMESPACE) {
            match local.as_ref() {
                b"type" => &mut xlink_type,
                b"href" => &mut href,
                b"actuate" => &mut actuate,
                _ => return invalid("unsupported XLink event-listener attribute"),
            }
        } else {
            return invalid("event-listener attribute has an unsupported namespace");
        };
        if slot.replace(value).is_some() {
            return invalid("duplicate script:event-listener attribute");
        }
    }

    let event_name = required(event_name, "script:event-name")?;
    let language = required(language, "script:language")?;
    validate_required_value(&event_name, "script:event-name")?;
    validate_required_value(&language, "script:language")?;
    let binding = match (macro_name, xlink_type, href, actuate) {
        (Some(value), None, None, None) => {
            validate_required_value(&value, "script:macro-name")?;
            ScriptBinding::MacroName(value)
        },
        (None, Some(kind), Some(href), actuate) if kind == "simple" => {
            if actuate.as_deref().is_some_and(|value| value != "onRequest") {
                return invalid("script event xlink:actuate must be 'onRequest'");
            }
            validate_required_value(&href, "xlink:href")?;
            ScriptBinding::Linked { href }
        },
        _ => {
            return invalid(
                "script:event-listener requires exactly script:macro-name or simple xlink:href",
            );
        },
    };
    Ok(ScriptEventListener {
        event_name,
        language,
        binding,
    })
}

fn required_script_language(
    reader: &ResolvedReader<'_>,
    element: &BytesStart<'_>,
) -> Result<String> {
    let mut language = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid office:script attribute: {error}"))
        })?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&namespace);
        if !bound_to(&namespace, SCRIPT_NAMESPACE) || local.as_ref() != b"language" {
            return invalid("office:script contains an unsupported attribute");
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| Error::InvalidFormat(format!("invalid script:language: {error}")))?
            .into_owned();
        if language.replace(value).is_some() {
            return invalid("office:script contains duplicate script:language");
        }
    }
    let language = required(language, "office:script requires script:language")?;
    validate_required_value(&language, "script:language")?;
    Ok(language)
}

fn capture_inner_xml(
    reader: &mut ResolvedReader<'_>,
    xml: &str,
    content_start: usize,
    expected_local: &[u8],
    limits: &XmlStreamLimits,
) -> Result<String> {
    let mut depth = 0usize;
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader.read_resolved_event().map_err(|error| {
            litchi_odf_common::map_resolved_reader_error(error, "invalid embedded script XML")
        })?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(_) => {
                depth += 1;
                if depth > MAX_XML_DEPTH.min(limits.max_depth()) {
                    return invalid(format!(
                        "embedded script XML exceeds the {MAX_XML_DEPTH} depth limit"
                    ));
                }
            },
            Event::End(element) if depth == 0 => {
                if !bound_to(&namespace, OFFICE_NAMESPACE)
                    || element.local_name().as_ref() != expected_local
                {
                    return invalid("mismatched office:script end element");
                }
                return xml
                    .get(content_start..event_start)
                    .map(ToOwned::to_owned)
                    .ok_or_else(|| {
                        Error::InvalidFormat("invalid embedded script range".to_string())
                    });
            },
            Event::End(_) => depth -= 1,
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in embedded scripts"),
            Event::Eof => return invalid("unterminated office:script element"),
            _ => {},
        }
    }
}

fn capture_full_xml(
    reader: &mut ResolvedReader<'_>,
    xml: &str,
    element_start: usize,
    expected_local: &[u8],
    limits: &XmlStreamLimits,
) -> Result<String> {
    let mut depth = 0usize;
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(|error| {
            litchi_odf_common::map_resolved_reader_error(error, "invalid preserved listener XML")
        })?;
        let namespace = namespace_kind(&namespace);
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(_) => {
                depth += 1;
                if depth > MAX_XML_DEPTH.min(limits.max_depth()) {
                    return invalid(format!(
                        "presentation listener exceeds the {MAX_XML_DEPTH} depth limit"
                    ));
                }
            },
            Event::End(element) if depth == 0 => {
                if !bound_to(&namespace, PRESENTATION_NAMESPACE)
                    || element.local_name().as_ref() != expected_local
                {
                    return invalid("mismatched presentation:event-listener end element");
                }
                return xml
                    .get(element_start..event_end)
                    .map(ToOwned::to_owned)
                    .ok_or_else(|| Error::InvalidFormat("invalid listener XML range".to_string()));
            },
            Event::End(_) => depth -= 1,
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in event listeners"),
            Event::Eof => return invalid("unterminated presentation:event-listener"),
            _ => {},
        }
    }
}

fn require_empty_element(
    reader: &mut ResolvedReader<'_>,
    expected_local: &[u8],
    expected_namespace: &[u8],
) -> Result<()> {
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(|error| {
            litchi_odf_common::map_resolved_reader_error(error, "invalid event-listener body")
        })?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::End(element)
                if bound_to(&namespace, expected_namespace)
                    && element.local_name().as_ref() == expected_local =>
            {
                return Ok(());
            },
            Event::Text(text) => {
                let value = text.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid listener text: {error}"))
                })?;
                if !value.trim().is_empty() {
                    return invalid("script:event-listener must be empty");
                }
            },
            Event::Comment(_) | Event::PI(_) => {},
            Event::Eof => return invalid("unterminated script:event-listener"),
            _ => return invalid("script:event-listener must not contain child elements"),
        }
    }
}

fn namespace_declarations(
    reader: &ResolvedReader<'_>,
    element: &BytesStart<'_>,
) -> Result<Vec<NamespaceDeclaration>> {
    let mut declarations: Vec<NamespaceDeclaration> = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid namespace declaration: {error}"))
        })?;
        let key = attribute.key.as_ref();
        let prefix = if key == b"xmlns" {
            Some(None)
        } else {
            key.strip_prefix(b"xmlns:")
                .map(|value| Some(String::from_utf8_lossy(value).into_owned()))
        };
        let Some(prefix) = prefix else { continue };
        let uri = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| Error::InvalidFormat(format!("invalid namespace URI: {error}")))?
            .into_owned();
        if uri.len() > MAX_VALUE_BYTES
            || prefix
                .as_ref()
                .is_some_and(|value| value.len() > MAX_VALUE_BYTES)
        {
            return invalid("namespace declaration exceeds the value-size limit");
        }
        if let Some(existing) = declarations.iter_mut().find(|item| item.prefix == prefix) {
            existing.uri = uri;
        } else {
            declarations.push(NamespaceDeclaration { prefix, uri });
        }
    }
    Ok(declarations)
}

fn merge_namespace_declarations(
    declarations: &mut Vec<NamespaceDeclaration>,
    additional: Vec<NamespaceDeclaration>,
) {
    for declaration in additional {
        if let Some(existing) = declarations
            .iter_mut()
            .find(|item| item.prefix == declaration.prefix)
        {
            existing.uri = declaration.uri;
        } else {
            declarations.push(declaration);
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct FragmentReport {
    pub(super) events: usize,
    pub(super) attributes: usize,
}

#[derive(Debug, Clone, Copy)]
pub(super) enum FragmentContainer<'a> {
    Script { language: &'a str },
    Presentation,
}

pub(super) fn validate_fragment(
    fragment: &str,
    declarations: &[NamespaceDeclaration],
    limits: &XmlStreamLimits,
    container: FragmentContainer<'_>,
) -> Result<FragmentReport> {
    // Keep the real office:scripts envelope here so namespace and depth
    // admission observes the same scopes and token boundaries as the writer.
    if fragment.len() > MAX_TEXT_BYTES
        || u64::try_from(fragment.len()).unwrap_or(u64::MAX) > limits.max_bytes()
    {
        return invalid(format!(
            "preserved script XML exceeds the {MAX_TEXT_BYTES} byte limit"
        ));
    }
    let capacity = fragment
        .len()
        .checked_add(512)
        .ok_or_else(|| Error::InvalidFormat("preserved script XML size overflow".to_string()))?;
    let mut wrapper = String::new();
    wrapper
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "preserved script XML validation",
            source,
        })?;
    append_scripts_open(&mut wrapper, declarations)?;
    wrapper.push('>');
    let wrapper_attributes = match container {
        FragmentContainer::Script { language } => {
            wrapper.push_str("<office:script script:language=\"");
            escape_attribute(&mut wrapper, language);
            wrapper.push_str("\">");
            wrapper.push_str(fragment);
            wrapper.push_str("</office:script>");
            1
        },
        FragmentContainer::Presentation => {
            wrapper.push_str("<office:event-listeners>");
            wrapper.push_str(fragment);
            wrapper.push_str("</office:event-listeners>");
            0
        },
    };
    wrapper.push_str("</office:scripts>");

    let mut reader = ResolvedReader::from_xml_document(&wrapper, *limits).map_err(|error| {
        litchi_odf_common::map_resolved_reader_error(error, "invalid preserved script XML")
    })?;
    let mut depth = 0usize;
    loop {
        match reader
            .read_resolved_event()
            .map_err(|error| {
                litchi_odf_common::map_resolved_reader_error(error, "invalid preserved script XML")
            })?
            .1
        {
            Event::Start(_) => {
                depth += 1;
                if depth > MAX_XML_DEPTH.min(limits.max_depth()) + 2 {
                    return invalid(format!(
                        "preserved script XML exceeds the {MAX_XML_DEPTH} depth limit"
                    ));
                }
            },
            Event::Empty(_) => {
                depth += 1;
                if depth > MAX_XML_DEPTH.min(limits.max_depth()) + 2 {
                    return invalid(format!(
                        "preserved script XML exceeds the {MAX_XML_DEPTH} depth limit"
                    ));
                }
                depth = depth.saturating_sub(1);
            },
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in script metadata"),
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(FragmentReport {
        events: reader.event_count().checked_sub(5).ok_or_else(|| {
            Error::InvalidFormat("preserved script XML report is incomplete".to_string())
        })?,
        attributes: reader
            .attribute_count()
            .checked_sub(emitted_namespace_count(declarations) + wrapper_attributes)
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "preserved script XML attribute report is incomplete".to_string(),
                )
            })?,
    })
}

pub(super) fn validate_static_output(
    declarations: &[NamespaceDeclaration],
    limits: &XmlStreamLimits,
) -> Result<()> {
    let mut output = String::new();
    append_scripts_open(&mut output, declarations)?;
    output.push('>');
    output.push_str("</office:scripts>");
    validate_generated_document(&output, limits)
}

pub(super) fn validate_empty_script(
    language: &str,
    declarations: &[NamespaceDeclaration],
    limits: &XmlStreamLimits,
) -> Result<()> {
    let mut output = String::new();
    append_scripts_open(&mut output, declarations)?;
    output.push('>');
    output.push_str("<office:script script:language=\"");
    escape_attribute(&mut output, language);
    output.push_str("\"/>");
    output.push_str("</office:scripts>");
    validate_generated_document(&output, limits)
}

pub(super) fn validate_script_listener(
    listener: &ScriptEventListener,
    declarations: &[NamespaceDeclaration],
    limits: &XmlStreamLimits,
) -> Result<()> {
    let mut output = String::new();
    append_scripts_open(&mut output, declarations)?;
    output.push('>');
    output.push_str("<office:event-listeners><script:event-listener script:event-name=\"");
    escape_attribute(&mut output, &listener.event_name);
    output.push_str("\" script:language=\"");
    escape_attribute(&mut output, &listener.language);
    match &listener.binding {
        ScriptBinding::MacroName(value) => {
            output.push_str("\" script:macro-name=\"");
            escape_attribute(&mut output, value);
            output.push_str("\"/>");
        },
        ScriptBinding::Linked { href } => {
            output.push_str("\" xlink:type=\"simple\" xlink:href=\"");
            escape_attribute(&mut output, href);
            output.push_str("\" xlink:actuate=\"onRequest\"/>");
        },
    }
    output.push_str("</office:event-listeners></office:scripts>");
    validate_generated_document(&output, limits)
}

fn validate_generated_document(output: &str, limits: &XmlStreamLimits) -> Result<()> {
    let mut reader = ResolvedReader::from_xml_document(output, *limits).map_err(|error| {
        litchi_odf_common::map_resolved_reader_error(error, "invalid emitted script XML")
    })?;
    loop {
        let (_, event) = reader.read_resolved_event().map_err(|error| {
            litchi_odf_common::map_resolved_reader_error(error, "invalid emitted script XML")
        })?;
        if matches!(event, Event::Eof) {
            break;
        }
    }
    Ok(())
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE_NAMESPACE => NamespaceKind::Office,
        ResolveResult::Bound(value) if value.as_ref() == SCRIPT_NAMESPACE => NamespaceKind::Script,
        ResolveResult::Bound(value) if value.as_ref() == PRESENTATION_NAMESPACE => {
            NamespaceKind::Presentation
        },
        ResolveResult::Bound(value) if value.as_ref() == XLINK_NAMESPACE => NamespaceKind::Xlink,
        _ => NamespaceKind::Other,
    }
}

fn bound_to(namespace: &NamespaceKind, expected: &[u8]) -> bool {
    match namespace {
        NamespaceKind::Office => expected == OFFICE_NAMESPACE,
        NamespaceKind::Script => expected == SCRIPT_NAMESPACE,
        NamespaceKind::Presentation => expected == PRESENTATION_NAMESPACE,
        NamespaceKind::Xlink => expected == XLINK_NAMESPACE,
        NamespaceKind::Other => false,
    }
}

pub(super) fn checked_text_bytes(current: usize, additional: usize) -> Result<usize> {
    let total = current
        .checked_add(additional)
        .ok_or_else(|| Error::InvalidFormat("script metadata size overflow".to_string()))?;
    if total > MAX_TEXT_BYTES {
        return invalid(format!(
            "script metadata exceeds the {MAX_TEXT_BYTES} aggregate byte limit"
        ));
    }
    Ok(total)
}

pub(super) fn validate_required_value(value: &str, name: &str) -> Result<()> {
    if value.is_empty() {
        return invalid(format!("{name} must not be empty"));
    }
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{name} exceeds the {MAX_VALUE_BYTES} byte limit"));
    }
    Ok(())
}

fn required(value: Option<String>, message: &str) -> Result<String> {
    value.ok_or_else(|| Error::InvalidFormat(message.to_string()))
}

fn ensure_listener_capacity(count: usize) -> Result<()> {
    if count >= MAX_LISTENER_COUNT {
        invalid(format!(
            "office:scripts exceeds the {MAX_LISTENER_COUNT} event-listener limit"
        ))
    } else {
        Ok(())
    }
}

fn script_listener_bytes(listener: &ScriptEventListener) -> usize {
    listener.event_name.len()
        + listener.language.len()
        + match &listener.binding {
            ScriptBinding::MacroName(value) => value.len(),
            ScriptBinding::Linked { href } => href.len(),
        }
}

fn is_namespace_declaration(key: &[u8]) -> bool {
    key == b"xmlns" || key.starts_with(b"xmlns:")
}

fn namespace_attribute(output: &mut String, prefix: Option<&str>, uri: &str) {
    output.push_str(" xmlns");
    if let Some(prefix) = prefix {
        output.push(':');
        output.push_str(prefix);
    }
    output.push_str("=\"");
    escape_attribute(output, uri);
    output.push('"');
}

fn escape_attribute(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '"' => output.push_str("&quot;"),
            '\r' => output.push_str("&#13;"),
            '\n' => output.push_str("&#10;"),
            '\t' => output.push_str("&#9;"),
            _ => output.push(character),
        }
    }
}
