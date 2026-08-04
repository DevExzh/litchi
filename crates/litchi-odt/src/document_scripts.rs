//! Inert ODF document script declarations.
//!
//! This module intentionally exposes script payloads as metadata only. It never
//! loads linked resources and never executes embedded script content.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::ResolveResult,
    reader::NsReader,
};

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const SCRIPT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const PRESENTATION_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const MAX_DOCUMENT_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_SCRIPT_COUNT: usize = 1_024;
const MAX_LISTENER_COUNT: usize = 4_096;
const MAX_TEXT_BYTES: usize = 8 * 1024 * 1024;
const MAX_VALUE_BYTES: usize = 64 * 1024;
const MAX_XML_DEPTH: usize = 128;

/// One inert `office:script` declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedScript {
    /// Required script language identifier.
    pub language: String,
    /// The exact inner XML payload. It is never interpreted or executed.
    pub content_xml: String,
}

/// The target stored by a `script:event-listener`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ScriptBinding {
    /// An inert macro name. Litchi never invokes it.
    MacroName(String),
    /// An inert linked script reference. Litchi never resolves it.
    Linked { href: String },
}

/// One typed `script:event-listener` declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ScriptEventListener {
    pub event_name: String,
    pub language: String,
    pub binding: ScriptBinding,
}

/// One child of the document-level `office:event-listeners` element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum EventListener {
    Script(ScriptEventListener),
    /// A presentation listener preserved as inert XML for lossless round trips.
    PresentationXml(String),
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct NamespaceDeclaration {
    prefix: Option<String>,
    uri: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Script,
    Presentation,
    Xlink,
    Other,
}

/// Semantic contents of an ODF `office:scripts` element.
#[derive(Debug, Clone, Default)]
pub struct Scripts {
    pub scripts: Vec<EmbeddedScript>,
    pub event_listeners: Vec<EventListener>,
    namespace_declarations: Vec<NamespaceDeclaration>,
}

impl PartialEq for Scripts {
    fn eq(&self, other: &Self) -> bool {
        self.scripts == other.scripts && self.event_listeners == other.event_listeners
    }
}

impl Eq for Scripts {}

impl Scripts {
    /// Validate resource limits, required values, and preserved XML fragments.
    pub fn validate(&self) -> Result<()> {
        if self.scripts.len() > MAX_SCRIPT_COUNT {
            return invalid(format!(
                "office:scripts exceeds the {MAX_SCRIPT_COUNT} script limit"
            ));
        }
        if self.event_listeners.len() > MAX_LISTENER_COUNT {
            return invalid(format!(
                "office:scripts exceeds the {MAX_LISTENER_COUNT} event-listener limit"
            ));
        }

        let mut text_bytes = 0usize;
        for script in &self.scripts {
            validate_required_value(&script.language, "script:language")?;
            text_bytes = checked_text_bytes(text_bytes, script.language.len())?;
            text_bytes = checked_text_bytes(text_bytes, script.content_xml.len())?;
            validate_fragment(&script.content_xml, &self.namespace_declarations)?;
        }

        for listener in &self.event_listeners {
            match listener {
                EventListener::Script(listener) => {
                    validate_required_value(&listener.event_name, "script:event-name")?;
                    validate_required_value(&listener.language, "script:language")?;
                    text_bytes = checked_text_bytes(text_bytes, listener.event_name.len())?;
                    text_bytes = checked_text_bytes(text_bytes, listener.language.len())?;
                    let value = match &listener.binding {
                        ScriptBinding::MacroName(value) => value,
                        ScriptBinding::Linked { href } => href,
                    };
                    validate_required_value(value, "script event target")?;
                    text_bytes = checked_text_bytes(text_bytes, value.len())?;
                },
                EventListener::PresentationXml(xml) => {
                    text_bytes = checked_text_bytes(text_bytes, xml.len())?;
                    validate_fragment(xml, &self.namespace_declarations)?;
                },
            }
        }
        Ok(())
    }

    /// Serialize a namespace-complete, deterministic `office:scripts` element.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(256 + self.total_payload_bytes());
        output.push_str("<office:scripts");
        namespace_attribute(
            &mut output,
            Some("office"),
            std::str::from_utf8(OFFICE_NAMESPACE).expect("ODF namespace is UTF-8"),
        );
        namespace_attribute(
            &mut output,
            Some("script"),
            std::str::from_utf8(SCRIPT_NAMESPACE).expect("ODF namespace is UTF-8"),
        );
        namespace_attribute(
            &mut output,
            Some("presentation"),
            std::str::from_utf8(PRESENTATION_NAMESPACE).expect("ODF namespace is UTF-8"),
        );
        namespace_attribute(
            &mut output,
            Some("xlink"),
            std::str::from_utf8(XLINK_NAMESPACE).expect("XLink namespace is UTF-8"),
        );
        for declaration in &self.namespace_declarations {
            if matches!(
                declaration.prefix.as_deref(),
                Some("office" | "script" | "presentation" | "xlink")
            ) {
                continue;
            }
            namespace_attribute(&mut output, declaration.prefix.as_deref(), &declaration.uri);
        }
        output.push('>');

        for script in &self.scripts {
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

        if !self.event_listeners.is_empty() {
            output.push_str("<office:event-listeners>");
            for listener in &self.event_listeners {
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
        Ok(output)
    }

    fn total_payload_bytes(&self) -> usize {
        self.scripts
            .iter()
            .map(|script| script.language.len() + script.content_xml.len())
            .chain(self.event_listeners.iter().map(|listener| match listener {
                EventListener::Script(listener) => {
                    listener.event_name.len()
                        + listener.language.len()
                        + match &listener.binding {
                            ScriptBinding::MacroName(value) => value.len(),
                            ScriptBinding::Linked { href } => href.len(),
                        }
                },
                EventListener::PresentationXml(xml) => xml.len(),
            }))
            .sum()
    }
}

/// Parse the optional direct `office:scripts` child from an ODF XML document.
pub fn parse_scripts(xml: &str) -> Result<Option<Scripts>> {
    if xml.len() > MAX_DOCUMENT_XML_BYTES {
        return invalid(format!(
            "ODF XML exceeds the {MAX_DOCUMENT_XML_BYTES} byte script-inventory limit"
        ));
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root_namespaces = Vec::new();
    let mut result = None;

    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODF script XML: {error}")))?;
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
                    )?);
                } else {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("ODF XML depth overflow".to_string())
                    })?;
                    if depth > MAX_XML_DEPTH {
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
        buffer.clear();
    }
    Ok(result)
}

fn parse_scripts_element(
    reader: &mut NsReader<&[u8]>,
    xml: &str,
    root_namespaces: &[NamespaceDeclaration],
    _content_start: usize,
) -> Result<Scripts> {
    let mut scripts = Vec::new();
    let mut event_listeners = Vec::new();
    let mut listener_container_seen = false;
    let mut text_bytes = 0usize;
    let mut buffer = Vec::new();

    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid office:scripts XML: {error}"))
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
                let content_xml = capture_inner_xml(reader, xml, event_end, b"script")?;
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
                event_listeners = parse_event_listeners(reader, xml, &mut text_bytes)?;
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
        buffer.clear();
    }

    Ok(Scripts {
        scripts,
        event_listeners,
        namespace_declarations: root_namespaces.to_vec(),
    })
}

fn parse_event_listeners(
    reader: &mut NsReader<&[u8]>,
    xml: &str,
    text_bytes: &mut usize,
) -> Result<Vec<EventListener>> {
    let mut listeners = Vec::new();
    let mut buffer = Vec::new();
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid event-listener XML: {error}"))
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
                let raw = capture_full_xml(reader, xml, event_start, b"event-listener")?;
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
        buffer.clear();
    }
    Ok(listeners)
}

fn parse_script_listener(
    reader: &NsReader<&[u8]>,
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

fn required_script_language(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<String> {
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
    reader: &mut NsReader<&[u8]>,
    xml: &str,
    content_start: usize,
    expected_local: &[u8],
) -> Result<String> {
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid embedded script XML: {error}"))
            })?;
        let namespace = namespace_kind(&namespace);
        match event {
            Event::Start(_) => {
                depth += 1;
                if depth > MAX_XML_DEPTH {
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
        buffer.clear();
    }
}

fn capture_full_xml(
    reader: &mut NsReader<&[u8]>,
    xml: &str,
    element_start: usize,
    expected_local: &[u8],
) -> Result<String> {
    let mut depth = 0usize;
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid preserved listener XML: {error}"))
            })?;
        let namespace = namespace_kind(&namespace);
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(_) => {
                depth += 1;
                if depth > MAX_XML_DEPTH {
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
        buffer.clear();
    }
}

fn require_empty_element(
    reader: &mut NsReader<&[u8]>,
    expected_local: &[u8],
    expected_namespace: &[u8],
) -> Result<()> {
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid event-listener body: {error}"))
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
        buffer.clear();
    }
}

fn namespace_declarations(
    reader: &NsReader<&[u8]>,
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

fn validate_fragment(fragment: &str, declarations: &[NamespaceDeclaration]) -> Result<()> {
    if fragment.len() > MAX_TEXT_BYTES {
        return invalid(format!(
            "preserved script XML exceeds the {MAX_TEXT_BYTES} byte limit"
        ));
    }
    let mut wrapper = String::with_capacity(fragment.len() + 256);
    wrapper.push_str("<litchi-root");
    for declaration in declarations {
        namespace_attribute(
            &mut wrapper,
            declaration.prefix.as_deref(),
            &declaration.uri,
        );
    }
    wrapper.push('>');
    wrapper.push_str(fragment);
    wrapper.push_str("</litchi-root>");

    let mut reader = NsReader::from_str(&wrapper);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buffer).map_err(|error| {
            Error::InvalidFormat(format!("invalid preserved script XML: {error}"))
        })? {
            Event::Start(_) => {
                depth += 1;
                if depth > MAX_XML_DEPTH + 1 {
                    return invalid(format!(
                        "preserved script XML exceeds the {MAX_XML_DEPTH} depth limit"
                    ));
                }
            },
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in script metadata"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
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
    matches!(
        (namespace, expected),
        (NamespaceKind::Office, OFFICE_NAMESPACE)
            | (NamespaceKind::Script, SCRIPT_NAMESPACE)
            | (NamespaceKind::Presentation, PRESENTATION_NAMESPACE)
            | (NamespaceKind::Xlink, XLINK_NAMESPACE)
    )
}

fn checked_text_bytes(current: usize, additional: usize) -> Result<usize> {
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

fn validate_required_value(value: &str, name: &str) -> Result<()> {
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

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const SCRIPT: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
    const PRESENTATION: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";

    #[test]
    fn parses_and_round_trips_inert_document_scripts() {
        // ODF 1.2/1.3 office:scripts grammar; LibreOffice xmlscripti.cxx consumes
        // the same office:script and office:event-listeners sequence.
        let xml = format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:s="{SCRIPT}" xmlns:p="{PRESENTATION}" xmlns:x="http://www.w3.org/1999/xlink" xmlns:ooo="http://openoffice.org/2004/office"><o:scripts><o:script s:language="ooo:Basic"><ooo:libraries><ooo:library name="A&amp;B"/></ooo:libraries></o:script><o:script s:language="Python"/><o:event-listeners><s:event-listener s:event-name="dom:load" s:language="ooo:script" s:macro-name="Standard.Module.Main"/><s:event-listener s:event-name="dom:unload" s:language="javascript" x:type="simple" x:href="Scripts/close.js" x:actuate="onRequest"/><p:event-listener s:event-name="dom:click" p:action="next-page"/></o:event-listeners></o:scripts><o:body/></o:document-content>"#
        );
        let scripts = parse_scripts(&xml).unwrap().unwrap();
        assert_eq!(scripts.scripts.len(), 2);
        assert_eq!(scripts.scripts[0].language, "ooo:Basic");
        assert!(scripts.scripts[0].content_xml.contains("ooo:library"));
        assert_eq!(scripts.event_listeners.len(), 3);
        assert!(matches!(
            &scripts.event_listeners[1],
            EventListener::Script(ScriptEventListener {
                binding: ScriptBinding::Linked { href },
                ..
            }) if href == "Scripts/close.js"
        ));

        let serialized = scripts.to_xml().unwrap();
        let reparsed = parse_scripts(&format!(
            r#"<office:document-content xmlns:office="{OFFICE}">{serialized}<office:body/></office:document-content>"#
        ))
        .unwrap()
        .unwrap();
        assert_eq!(reparsed, scripts);
    }

    #[test]
    fn rejects_active_or_malformed_script_metadata() {
        let wrap = |body: &str| {
            format!(
                r#"<o:document-content xmlns:o="{OFFICE}" xmlns:s="{SCRIPT}" xmlns:x="http://www.w3.org/1999/xlink"><o:scripts>{body}</o:scripts><o:body/></o:document-content>"#
            )
        };
        for body in [
            r#"<o:script/>"#,
            r#"<o:event-listeners/><o:script s:language="Python"/>"#,
            r#"<o:event-listeners><s:event-listener s:event-name="load" s:language="x"/></o:event-listeners>"#,
            r#"<o:event-listeners><s:event-listener s:event-name="load" s:language="x" s:macro-name="M" x:type="simple" x:href="S"/></o:event-listeners>"#,
            r#"<o:event-listeners><s:event-listener s:event-name="load" s:language="x" x:type="extended" x:href="S"/></o:event-listeners>"#,
            r#"<o:script s:language="Python"><!DOCTYPE x></o:script>"#,
        ] {
            assert!(parse_scripts(&wrap(body)).is_err(), "{body}");
        }
    }

    #[test]
    fn rejects_nested_and_duplicate_script_containers() {
        let nested = format!(
            r#"<o:document-content xmlns:o="{OFFICE}"><o:body><o:scripts/></o:body></o:document-content>"#
        );
        assert!(parse_scripts(&nested).is_err());
        let duplicate = format!(
            r#"<o:document-content xmlns:o="{OFFICE}"><o:scripts/><o:scripts/><o:body/></o:document-content>"#
        );
        assert!(parse_scripts(&duplicate).is_err());
    }
}
