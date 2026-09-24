//! Bounded `SpreadsheetML` connection-part XML codec.

use super::invalid;
use super::model::{
    CORE_NAMESPACE, Connection, ConnectionParameter, Connections, CredentialsMethod,
    DatabaseProperties, HtmlFormatting, MAX_CONNECTIONS, MAX_DOM_DEPTH, MAX_DOM_NODES,
    MAX_PARAMETERS, MAX_STRING_BYTES, MAX_TEXT_FIELDS, MAX_WEB_TABLES, MAX_XML_BYTES,
    OlapProperties, ParameterType, STRICT_NAMESPACE, TextField, TextFieldType, TextFileType,
    TextImportProperties, TextQualifier, WebQueryProperties, WebTableSelector, validate,
};
use super::namespace::{NamespaceContext, NamespaceDecl, NamespaceLimits};
use crate::error::{Error as XlsxError, Result as XlsxResult};
use litchi_core::sheet::Result;
use litchi_core::xml::ReaderOrigin;
use quick_xml::{
    Reader, XmlVersion,
    encoding::Decoder,
    events::{BytesStart, Event},
};
use std::{
    borrow::Cow,
    collections::{HashMap, HashSet},
    hash::{Hash, Hasher},
    ops::Range,
    sync::Arc,
};

type NamespaceUri = Arc<str>;

fn namespace_limits() -> NamespaceLimits {
    NamespaceLimits::new(MAX_DOM_DEPTH, MAX_DOM_NODES, MAX_XML_BYTES)
}

fn namespace_root() -> NamespaceContext {
    NamespaceContext::root(namespace_limits())
}

fn namespace_from_bindings(bindings: &[(String, String)]) -> Result<NamespaceContext> {
    namespace_root().child(
        bindings
            .iter()
            .map(|(prefix, uri)| NamespaceDecl::new(prefix, uri)),
    )
}

#[derive(Clone)]
struct Attr {
    q: String,
    ns: NamespaceUri,
    l: String,
    v: String,
}
#[derive(Clone)]
enum Content {
    Node(Node),
    Text(String),
    CData(String),
    Comment(String),
}
#[derive(Clone)]
pub(super) struct Node {
    q: String,
    ns: NamespaceUri,
    l: String,
    attrs: Vec<Attr>,
    context: NamespaceContext,
    content: Vec<Content>,
}

impl Connections {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_XML_BYTES {
            return Err(invalid("connections part exceeds 16 MiB"));
        }
        // The MCE processor rejects otherwise legal processing instructions.
        // They do not contribute to the typed connection projection, while
        // source-bound callers retain the original bytes for exact splicing.
        // Remove only PI events before MCE processing; all schema-bearing
        // markup still goes through the normal full parser and validator.
        let without_pi = strip_processing_instructions(xml)?;
        let x = litchi_ooxml_common::mce::process_ooxml(without_pi.as_ref())?;
        if x.len() > MAX_XML_BYTES {
            return Err(invalid("processed connections part exceeds 16 MiB"));
        }
        project(&parse_dom(x.as_ref())?)
    }

    /// Parse one connections part under a caller-selected byte ceiling.
    ///
    /// The normal public parser retains its historical fixed profile.  The
    /// Custom Data owner uses this bounded entry point for graph admission so
    /// MCE preprocessing and the typed DOM never see a member larger than the
    /// host profile or the immutable 16 MiB parser ceiling.
    pub(super) fn parse_with_limits(
        xml: &[u8],
        max_bytes: usize,
        max_nodes: usize,
        max_events: usize,
        max_depth: usize,
        max_string_bytes: usize,
        max_namespace_bytes: usize,
        max_attributes: usize,
        max_temporary_bytes: usize,
    ) -> XlsxResult<Self> {
        let maximum = max_bytes.min(MAX_XML_BYTES);
        let nodes = max_nodes.min(MAX_DOM_NODES);
        // Preserve a caller's zero-event profile. The first lexical event
        // must be refused rather than silently widening the request.
        let events = max_events;
        let depth = max_depth.min(MAX_DOM_DEPTH);
        if xml.len() > maximum {
            return Err(connection_limit(
                xml.len(),
                maximum,
                "connections input XML bytes",
            ));
        }
        preflight_xml_with_limits(
            xml,
            nodes,
            events,
            depth,
            max_string_bytes,
            max_namespace_bytes,
            max_attributes,
        )?;
        let without_pi = strip_processing_instructions_with_limit(xml, max_temporary_bytes)?;
        let mce_output_limit = if contains_mce_markup(without_pi.as_ref()) {
            maximum.min(max_temporary_bytes)
        } else {
            maximum
        };
        let processed = litchi_ooxml_common::mce::process_markup_compatibility(
            without_pi.as_ref(),
            &litchi_ooxml_common::mce::Capabilities::default(),
            &litchi_ooxml_common::mce::Limits {
                max_input_bytes: maximum,
                max_output_bytes: mce_output_limit,
                max_depth: depth,
                max_namespace_bindings: max_attributes.min(4096),
                max_directive_tokens: events.min(4096),
                max_choices_per_alternate: events.min(1024),
            },
        )
        .map_err(|error| {
            let mapped = map_mce_error(
                error,
                maximum,
                mce_output_limit,
                depth,
                max_attributes.min(4096),
                events.min(4096),
                events.min(1024),
                "connections",
            );
            if mce_output_limit < maximum
                && matches!(
                    mapped,
                    XlsxError::ResourceLimit(litchi_core::ResourceLimit {
                        resource: litchi_core::Resource::OutputBytes,
                        ..
                    })
                )
            {
                temporary_limit(
                    mce_output_limit.saturating_add(1),
                    mce_output_limit,
                    "connections MCE output bytes",
                )
            } else {
                mapped
            }
        })?;
        if processed.xml.len() > maximum {
            return Err(xml_resource_limit(
                litchi_core::Resource::OutputBytes,
                processed.xml.len(),
                maximum,
                "connections processed XML bytes",
            ));
        }
        let root = parse_dom_with_limits(
            processed.xml.as_ref(),
            nodes,
            events,
            depth,
            max_string_bytes,
            max_namespace_bytes,
            max_attributes,
        )
        .map_err(|error| XlsxError::Invalid(error.to_string()))?;
        project(&root).map_err(|error| XlsxError::Invalid(error.to_string()))
    }
    pub fn to_xml(&self, strict: bool) -> Result<Vec<u8>> {
        validate(self)?;
        let ns = if strict {
            STRICT_NAMESPACE
        } else {
            CORE_NAMESPACE
        };
        let mut x = BoundedXml::new();
        x.push_str(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><connections xmlns=\"",
        )?;
        x.push_str(ns)?;
        x.push_str("\">")?;
        for c in &self.connections {
            write_connection(&mut x, c, strict)?;
        }
        x.push_str("</connections>")?;
        Ok(x.finish())
    }
}

fn connection_limit(observed: usize, maximum: usize, name: &'static str) -> XlsxError {
    XlsxError::ResourceLimit(litchi_core::ResourceLimit {
        resource: litchi_core::Resource::InputBytes,
        observed: observed as u64,
        limit: maximum as u64,
        scope: Arc::<str>::from(format!("XLSX Custom Data {name}")),
    })
}

fn temporary_limit(observed: usize, maximum: usize, name: &'static str) -> XlsxError {
    XlsxError::ResourceLimit(litchi_core::ResourceLimit {
        resource: litchi_core::Resource::Memory,
        observed: observed as u64,
        limit: maximum as u64,
        scope: Arc::<str>::from(format!("XLSX Custom Data {name}")),
    })
}

fn xml_resource_limit(
    resource: litchi_core::Resource,
    observed: usize,
    maximum: usize,
    name: &'static str,
) -> XlsxError {
    XlsxError::ResourceLimit(litchi_core::ResourceLimit {
        resource,
        observed: observed as u64,
        limit: maximum as u64,
        scope: Arc::<str>::from(format!("XLSX Custom Data {name}")),
    })
}

pub(super) fn map_mce_error(
    error: litchi_ooxml_common::mce::Error,
    input_bytes: usize,
    output_bytes: usize,
    depth: usize,
    namespace_bindings: usize,
    directive_tokens: usize,
    choices: usize,
    scope: &'static str,
) -> XlsxError {
    let litchi_ooxml_common::mce::Error::LimitExceeded(resource) = &error else {
        return XlsxError::MarkupCompatibility(error);
    };
    let (kind, maximum) = match resource.as_str() {
        "input bytes" => (litchi_core::Resource::InputBytes, input_bytes),
        "output bytes" => (litchi_core::Resource::OutputBytes, output_bytes),
        "depth" => (litchi_core::Resource::Depth, depth),
        "namespace bindings" => (litchi_core::Resource::Objects, namespace_bindings),
        "directive tokens" => (litchi_core::Resource::Objects, directive_tokens),
        "choices" => (litchi_core::Resource::Objects, choices),
        _ => return XlsxError::MarkupCompatibility(error),
    };
    XlsxError::ResourceLimit(litchi_core::ResourceLimit {
        resource: kind,
        observed: maximum.saturating_add(1) as u64,
        limit: maximum as u64,
        scope: Arc::<str>::from(format!("XLSX Custom Data {scope} MCE {resource}")),
    })
}

/// Admit the borrowed lexical shape before any PI copy, MCE expansion, or
/// namespace resolver is allowed to allocate state. The DOM parser repeats
/// these checks while constructing its owned projection.
pub(super) fn preflight_xml_with_limits(
    xml: &[u8],
    max_nodes: usize,
    max_events: usize,
    max_depth: usize,
    max_string_bytes: usize,
    max_namespace_bytes: usize,
    max_attributes: usize,
) -> XlsxResult<()> {
    std::str::from_utf8(xml).map_err(|error| XlsxError::Invalid(error.to_string()))?;
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut stack = Vec::<usize>::new();
    let mut nodes = 0usize;
    let mut events = 0usize;
    let mut namespace_bytes = 0usize;
    loop {
        events = events.checked_add(1).ok_or_else(|| {
            xml_resource_limit(
                litchi_core::Resource::Objects,
                usize::MAX,
                max_events,
                "XML events",
            )
        })?;
        if events > max_events {
            return Err(xml_resource_limit(
                litchi_core::Resource::Objects,
                events,
                max_events,
                "XML events",
            ));
        }
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    xml_resource_limit(
                        litchi_core::Resource::Objects,
                        usize::MAX,
                        max_nodes,
                        "XML nodes",
                    )
                })?;
                if nodes > max_nodes {
                    return Err(xml_resource_limit(
                        litchi_core::Resource::Objects,
                        nodes,
                        max_nodes,
                        "XML nodes",
                    ));
                }
                let prospective_depth = stack.len().saturating_add(1);
                if prospective_depth > max_depth {
                    return Err(xml_resource_limit(
                        litchi_core::Resource::Depth,
                        prospective_depth,
                        max_depth,
                        "XML depth",
                    ));
                }
                let local_namespace_bytes = preflight_element(
                    &element,
                    namespace_bytes,
                    max_string_bytes,
                    max_namespace_bytes,
                    max_attributes,
                )?;
                namespace_bytes = namespace_bytes
                    .checked_add(local_namespace_bytes)
                    .ok_or_else(|| {
                        xml_resource_limit(
                            litchi_core::Resource::InputBytes,
                            usize::MAX,
                            max_namespace_bytes,
                            "XML namespace bytes",
                        )
                    })?;
                stack.push(local_namespace_bytes);
            },
            Ok(Event::Empty(element)) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    xml_resource_limit(
                        litchi_core::Resource::Objects,
                        usize::MAX,
                        max_nodes,
                        "XML nodes",
                    )
                })?;
                if nodes > max_nodes {
                    return Err(xml_resource_limit(
                        litchi_core::Resource::Objects,
                        nodes,
                        max_nodes,
                        "XML nodes",
                    ));
                }
                let prospective_depth = stack.len().saturating_add(1);
                if prospective_depth > max_depth {
                    return Err(xml_resource_limit(
                        litchi_core::Resource::Depth,
                        prospective_depth,
                        max_depth,
                        "XML depth",
                    ));
                }
                preflight_element(
                    &element,
                    namespace_bytes,
                    max_string_bytes,
                    max_namespace_bytes,
                    max_attributes,
                )?;
            },
            Ok(Event::End(_)) => {
                let local = stack
                    .pop()
                    .ok_or_else(|| XlsxError::Invalid("unexpected closing element".into()))?;
                namespace_bytes = namespace_bytes.checked_sub(local).ok_or_else(|| {
                    XlsxError::Invalid("XML namespace byte count underflow".into())
                })?;
            },
            Ok(Event::Text(text)) => preflight_payload(text.as_ref(), max_string_bytes)?,
            Ok(Event::CData(text)) => preflight_payload(text.as_ref(), max_string_bytes)?,
            Ok(Event::Comment(text)) => preflight_payload(text.as_ref(), max_string_bytes)?,
            Ok(Event::PI(text)) => preflight_payload(text.as_ref(), max_string_bytes)?,
            Ok(Event::GeneralRef(text)) => preflight_payload(text.as_ref(), max_string_bytes)?,
            Ok(Event::DocType(_)) => return Err(XlsxError::Invalid("DTDs are rejected".into())),
            Ok(Event::Decl(_)) => {},
            Ok(Event::Eof) => break,
            Err(error) => return Err(XlsxError::Invalid(xml_error(error).to_string())),
        }
    }
    if !stack.is_empty() {
        return Err(XlsxError::Invalid("unterminated XML".into()));
    }
    Ok(())
}

pub(super) fn contains_mce_markup(xml: &[u8]) -> bool {
    const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
    xml.windows(MCE_NAMESPACE.len())
        .any(|window| window == MCE_NAMESPACE)
}

fn preflight_payload(bytes: &[u8], maximum: usize) -> XlsxResult<()> {
    if bytes.len() > maximum {
        return Err(xml_resource_limit(
            litchi_core::Resource::InputBytes,
            bytes.len(),
            maximum,
            "XML string bytes",
        ));
    }
    Ok(())
}

fn preflight_element(
    element: &BytesStart<'_>,
    inherited_namespace_bytes: usize,
    max_string_bytes: usize,
    max_namespace_bytes: usize,
    max_attributes: usize,
) -> XlsxResult<usize> {
    if element.name().as_ref().len() > max_string_bytes {
        return Err(xml_resource_limit(
            litchi_core::Resource::InputBytes,
            element.name().as_ref().len(),
            max_string_bytes,
            "XML element name bytes",
        ));
    }
    let mut local_namespace_bytes = 0usize;
    let mut attribute_count = 0usize;
    for attribute in element.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| XlsxError::Invalid(error.to_string()))?;
        attribute_count = attribute_count.checked_add(1).ok_or_else(|| {
            xml_resource_limit(
                litchi_core::Resource::Objects,
                usize::MAX,
                max_attributes,
                "XML attributes",
            )
        })?;
        if attribute_count > max_attributes {
            return Err(xml_resource_limit(
                litchi_core::Resource::Objects,
                attribute_count,
                max_attributes,
                "XML attributes",
            ));
        }
        if attribute.key.as_ref().len() > max_string_bytes {
            return Err(xml_resource_limit(
                litchi_core::Resource::InputBytes,
                attribute.key.as_ref().len(),
                max_string_bytes,
                "XML attribute name bytes",
            ));
        }
        if attribute.value.as_ref().len() > max_string_bytes {
            return Err(xml_resource_limit(
                litchi_core::Resource::InputBytes,
                attribute.value.as_ref().len(),
                max_string_bytes,
                "XML attribute value bytes",
            ));
        }
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            let prefix_bytes = key.strip_prefix(b"xmlns:").map_or(0, |prefix| prefix.len());
            local_namespace_bytes = local_namespace_bytes
                .checked_add(prefix_bytes)
                .and_then(|value| value.checked_add(attribute.value.as_ref().len()))
                .ok_or_else(|| {
                    xml_resource_limit(
                        litchi_core::Resource::InputBytes,
                        usize::MAX,
                        max_namespace_bytes,
                        "XML namespace bytes",
                    )
                })?;
        }
    }
    let total = inherited_namespace_bytes
        .checked_add(local_namespace_bytes)
        .ok_or_else(|| {
            xml_resource_limit(
                litchi_core::Resource::InputBytes,
                usize::MAX,
                max_namespace_bytes,
                "XML namespace bytes",
            )
        })?;
    if total > max_namespace_bytes {
        return Err(xml_resource_limit(
            litchi_core::Resource::InputBytes,
            total,
            max_namespace_bytes,
            "XML namespace bytes",
        ));
    }
    Ok(local_namespace_bytes)
}

fn strip_processing_instructions<'a>(xml: &'a [u8]) -> Result<Cow<'a, [u8]>> {
    let mut reader = Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let position = |reader: &Reader<&[u8]>| {
        origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("processing-instruction source offset exceeds usize"))
    };
    let mut output: Option<Vec<u8>> = None;
    let mut cursor = 0;
    loop {
        let start = position(&reader)?;
        match reader.read_event() {
            Ok(Event::PI(_)) => {
                let end = position(&reader)?;
                if start < cursor || end < start || end > xml.len() {
                    return Err(invalid("invalid processing-instruction source span"));
                }
                if output.is_none() {
                    let mut bytes = Vec::new();
                    // The output only removes source bytes. Reserve one bounded
                    // buffer without retaining one range per PI event.
                    bytes
                        .try_reserve_exact(xml.len() - (end - start))
                        .map_err(|_| invalid("connections PI-filter output allocation failed"))?;
                    output = Some(bytes);
                }
                if let Some(bytes) = output.as_mut() {
                    bytes.extend_from_slice(&xml[cursor..start]);
                }
                cursor = end;
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(xml_error(error)),
        }
    }
    match output {
        Some(mut bytes) => {
            bytes.extend_from_slice(&xml[cursor..]);
            Ok(Cow::Owned(bytes))
        },
        None => Ok(Cow::Borrowed(xml)),
    }
}

/// Strip inert processing instructions while bounding the owned intermediate.
pub(super) fn strip_processing_instructions_with_limit<'a>(
    xml: &'a [u8],
    max_temporary_bytes: usize,
) -> XlsxResult<Cow<'a, [u8]>> {
    let mut reader = Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let position = |reader: &Reader<&[u8]>| {
        origin.offset(reader.buffer_position()).ok_or_else(|| {
            XlsxError::Invalid("processing-instruction source offset exceeds usize".into())
        })
    };
    let mut output: Option<Vec<u8>> = None;
    let mut cursor = 0usize;
    loop {
        let start = position(&reader)?;
        match reader.read_event() {
            Ok(Event::PI(_)) => {
                let end = position(&reader)?;
                if start < cursor || end < start || end > xml.len() {
                    return Err(XlsxError::Invalid(
                        "invalid processing-instruction source span".into(),
                    ));
                }
                let bytes = output.get_or_insert_with(Vec::new);
                let segment = &xml[cursor..start];
                let output_len = bytes.len().checked_add(segment.len()).ok_or_else(|| {
                    temporary_limit(
                        usize::MAX,
                        max_temporary_bytes,
                        "connections PI-filter bytes",
                    )
                })?;
                if output_len > max_temporary_bytes {
                    return Err(temporary_limit(
                        output_len,
                        max_temporary_bytes,
                        "connections PI-filter bytes",
                    ));
                }
                bytes
                    .try_reserve_exact(segment.len())
                    .map_err(|source| XlsxError::Allocation {
                        resource: "Custom Data connections PI-filter output",
                        source,
                    })?;
                bytes.extend_from_slice(segment);
                cursor = end;
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(XlsxError::Invalid(xml_error(error).to_string())),
        }
    }
    let Some(mut bytes) = output else {
        return Ok(Cow::Borrowed(xml));
    };
    let segment = &xml[cursor..];
    let output_len = bytes.len().checked_add(segment.len()).ok_or_else(|| {
        temporary_limit(
            usize::MAX,
            max_temporary_bytes,
            "connections PI-filter bytes",
        )
    })?;
    if output_len > max_temporary_bytes {
        return Err(temporary_limit(
            output_len,
            max_temporary_bytes,
            "connections PI-filter bytes",
        ));
    }
    bytes
        .try_reserve_exact(segment.len())
        .map_err(|source| XlsxError::Allocation {
            resource: "Custom Data connections PI-filter output",
            source,
        })?;
    bytes.extend_from_slice(segment);
    Ok(Cow::Owned(bytes))
}

pub(super) fn parse_dom(xml: &[u8]) -> Result<Node> {
    parse_dom_with_limits(
        xml,
        MAX_DOM_NODES,
        MAX_DOM_NODES.saturating_mul(4),
        MAX_DOM_DEPTH,
        MAX_STRING_BYTES,
        MAX_XML_BYTES,
        MAX_DOM_NODES,
    )
}

pub(super) fn parse_dom_with_limits(
    xml: &[u8],
    max_nodes: usize,
    max_events: usize,
    max_depth: usize,
    max_string_bytes: usize,
    max_namespace_bytes: usize,
    max_attributes: usize,
) -> Result<Node> {
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut rd = Reader::from_reader(xml);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut count = 0;
    let mut events = 0usize;
    loop {
        let d = rd.decoder();
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("connections XML event count overflow"))?;
        if events > max_events {
            return Err(invalid("connections XML event limit exceeded"));
        }
        match rd.read_event() {
            Ok(Event::Start(e)) => {
                count += 1;
                if count > max_nodes || stack.len().saturating_add(1) > max_depth {
                    return Err(invalid("connections XML resource limit exceeded"));
                }
                stack
                    .try_reserve(1)
                    .map_err(|_source| invalid("connections XML stack allocation failed"))?;
                stack.push(make(
                    &e,
                    d,
                    &stack,
                    max_depth,
                    max_namespace_bytes,
                    max_attributes,
                    max_string_bytes,
                )?);
            },
            Ok(Event::Empty(e)) => {
                count += 1;
                if count > max_nodes || stack.len().saturating_add(1) > max_depth {
                    return Err(invalid("connections node limit exceeded"));
                }
                let n = make(
                    &e,
                    d,
                    &stack,
                    max_depth,
                    max_namespace_bytes,
                    max_attributes,
                    max_string_bytes,
                )?;
                attach(&mut stack, &mut root, n)?;
            },
            Ok(Event::End(_)) => {
                let n = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element"))?;
                attach(&mut stack, &mut root, n)?;
            },
            Ok(Event::Text(t)) => {
                if t.as_ref().len() > max_string_bytes {
                    return Err(invalid("connections XML string limit exceeded"));
                }
                let v = t.decode().map_err(xml_error)?.into_owned();
                if v.len() > max_string_bytes {
                    return Err(invalid("connections XML string limit exceeded"));
                }
                if let Some(n) = stack.last_mut() {
                    n.content
                        .try_reserve(1)
                        .map_err(|_source| invalid("connections XML content allocation failed"))?;
                    n.content.push(Content::Text(v));
                } else if !v.trim().is_empty() {
                    return Err(invalid("text outside connections"));
                }
            },
            Ok(Event::CData(t)) => {
                if t.as_ref().len() > max_string_bytes {
                    return Err(invalid("connections XML string limit exceeded"));
                }
                if let Some(n) = stack.last_mut() {
                    n.content
                        .try_reserve(1)
                        .map_err(|_source| invalid("connections XML content allocation failed"))?;
                    n.content
                        .push(Content::CData(t.decode().map_err(xml_error)?.into_owned()));
                } else {
                    return Err(invalid("CDATA outside connections"));
                }
            },
            Ok(Event::Comment(t)) => {
                if t.as_ref().len() > max_string_bytes {
                    return Err(invalid("connections XML string limit exceeded"));
                }
                if let Some(n) = stack.last_mut() {
                    n.content
                        .try_reserve(1)
                        .map_err(|_source| invalid("connections XML content allocation failed"))?;
                    n.content.push(Content::Comment(
                        t.decode().map_err(xml_error)?.into_owned(),
                    ));
                }
            },
            Ok(Event::GeneralRef(t)) => {
                if t.as_ref().len() > max_string_bytes {
                    return Err(invalid("connections XML string limit exceeded"));
                }
                let value = litchi_ooxml_common::xml::decode_xml_reference(&t)?;
                if value.len() > max_string_bytes {
                    return Err(invalid("connections XML string limit exceeded"));
                }
                if let Some(n) = stack.last_mut() {
                    n.content
                        .try_reserve(1)
                        .map_err(|_source| invalid("connections XML content allocation failed"))?;
                    n.content.push(Content::Text(value));
                } else {
                    return Err(invalid("entity outside connections"));
                }
            },
            Ok(Event::DocType(_)) => {
                return Err(invalid("DTDs are rejected"));
            },
            // Processing instructions are legal XML around and inside a
            // connections part. They are intentionally omitted from the
            // typed projection; source-bound callers retain the original XML
            // and splice only the attributes they own.
            Ok(Event::PI(_)) => {},
            Ok(Event::Decl(_)) => {},
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(e)),
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated connections XML"));
    }
    root.ok_or_else(|| invalid("missing connections root"))
}
fn make(
    e: &BytesStart<'_>,
    d: Decoder,
    stack: &[Node],
    max_depth: usize,
    max_namespace_bytes: usize,
    max_attributes: usize,
    max_string_bytes: usize,
) -> Result<Node> {
    let raw_name = e.name();
    if raw_name.as_ref().len() > max_string_bytes {
        return Err(invalid("connections XML string limit exceeded"));
    }
    let q = std::str::from_utf8(raw_name.as_ref())
        .map_err(xml_error)?
        .to_string();
    let namespace_limits = NamespaceLimits::new(max_depth, max_attributes, max_namespace_bytes);
    let mut context = stack
        .last()
        .map(|node| node.context.clone())
        .unwrap_or_else(|| NamespaceContext::root(namespace_limits));
    let mut raw = Vec::new();
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xml_error)?;
        if raw.len() >= max_attributes {
            return Err(invalid("connections XML attribute limit exceeded"));
        }
        if a.value.as_ref().len() > max_string_bytes {
            return Err(invalid("connections XML string limit exceeded"));
        }
        if a.key.as_ref().len() > max_string_bytes {
            return Err(invalid("connections XML string limit exceeded"));
        }
        let qualified = std::str::from_utf8(a.key.as_ref()).map_err(xml_error)?;
        let value = a
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, d)
            .map_err(xml_error)?;
        if value.as_ref().len() > max_string_bytes {
            return Err(invalid("connections XML string limit exceeded"));
        }
        if qualified == "xmlns" || qualified.starts_with("xmlns:") {
            let prefix = qualified.strip_prefix("xmlns:").unwrap_or("");
            let namespace_bytes = prefix
                .len()
                .checked_add(value.len())
                .ok_or_else(|| invalid("connections XML namespace byte count overflow"))?;
            if namespace_bytes > max_namespace_bytes {
                return Err(invalid("connections XML namespace byte limit exceeded"));
            }
        }
        raw.try_reserve(1)
            .map_err(|_source| invalid("connections XML attribute allocation failed"))?;
        raw.push((qualified.to_owned(), value.into_owned()));
    }
    context = context.child(raw.iter().filter_map(|(qualified, value)| {
        if qualified == "xmlns" || qualified.starts_with("xmlns:") {
            Some(NamespaceDecl::new(
                qualified.strip_prefix("xmlns:").unwrap_or(""),
                value,
            ))
        } else {
            None
        }
    }))?;
    let (pr, lo) = split(&q)?;
    if lo.len() > max_string_bytes {
        return Err(invalid("connections XML string limit exceeded"));
    }
    let local = lo.to_string();
    let ns = resolve(&context, pr)?;
    let mut attrs = Vec::new();
    attrs
        .try_reserve(raw.len())
        .map_err(|_source| invalid("connections XML attribute allocation failed"))?;
    for (q, v) in raw {
        if q == "xmlns" || q.starts_with("xmlns:") {
            continue;
        }
        let (pr, lo) = split(&q)?;
        if lo.len() > max_string_bytes {
            return Err(invalid("connections XML string limit exceeded"));
        }
        let ans = if pr.is_empty() {
            Arc::from("")
        } else {
            resolve(&context, pr)?
        };
        let local = lo.to_string();
        attrs.push(Attr {
            q,
            ns: ans,
            l: local,
            v,
        });
    }
    Ok(Node {
        q,
        ns,
        l: local,
        attrs,
        context,
        content: Vec::new(),
    })
}
fn attach(stack: &mut [Node], root: &mut Option<Node>, n: Node) -> Result<()> {
    if let Some(p) = stack.last_mut() {
        p.content
            .try_reserve(1)
            .map_err(|_source| invalid("connections XML content allocation failed"))?;
        p.content.push(Content::Node(n));
    } else if root.replace(n).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

fn project(n: &Node) -> Result<Connections> {
    expect(n, "connections")?;
    noattrs(n)?;
    let mut out = Vec::new();
    for c in kids(n)? {
        // Future producers may place extension elements beside the typed
        // connection catalog. They stay in the source snapshot and are not
        // projected into the semantic model.
        if c.ns.as_ref() != CORE_NAMESPACE && c.ns.as_ref() != STRICT_NAMESPACE
            || c.l != "connection"
        {
            continue;
        }
        if out.len() >= MAX_CONNECTIONS {
            return Err(invalid("connection limit exceeded"));
        }
        out.push(parse_connection(c)?);
    }
    if out.is_empty() {
        return Err(invalid("connections requires at least one connection"));
    }
    let value = Connections { connections: out };
    validate(&value)?;
    Ok(value)
}
fn parse_connection(n: &Node) -> Result<Connection> {
    expect(n, "connection")?;
    let mut c = Connection {
        id: u32req(n, "id")?,
        source_file: aopt(n, "sourceFile")?,
        odc_file: aopt(n, "odcFile")?,
        keep_alive: bopt(n, "keepAlive")?,
        interval: u32opt(n, "interval")?,
        name: aopt(n, "name")?,
        description: aopt(n, "description")?,
        connection_type: u32opt(n, "type")?,
        reconnection_method: u32opt(n, "reconnectionMethod")?,
        refreshed_version: u8req(n, "refreshedVersion")?,
        min_refreshable_version: u8opt(n, "minRefreshableVersion")?,
        save_password: bopt(n, "savePassword")?,
        new_connection: bopt(n, "new")?,
        deleted: bopt(n, "deleted")?,
        only_use_connection_file: bopt(n, "onlyUseConnectionFile")?,
        background: bopt(n, "background")?,
        refresh_on_load: bopt(n, "refreshOnLoad")?,
        save_data: bopt(n, "saveData")?,
        credentials: aopt(n, "credentials")?.map(parse_credentials).transpose()?,
        single_sign_on_id: aopt(n, "singleSignOnId")?,
        database: None,
        olap: None,
        web: None,
        text: None,
        parameters: None,
        extension_xml: None,
    };
    only(
        n,
        &[
            "id",
            "sourceFile",
            "odcFile",
            "keepAlive",
            "interval",
            "name",
            "description",
            "type",
            "reconnectionMethod",
            "refreshedVersion",
            "minRefreshableVersion",
            "savePassword",
            "new",
            "deleted",
            "onlyUseConnectionFile",
            "background",
            "refreshOnLoad",
            "saveData",
            "credentials",
            "singleSignOnId",
        ],
    )?;
    let mut order = 0;
    for child in kids(n)? {
        if child.ns.as_ref() != CORE_NAMESPACE && child.ns.as_ref() != STRICT_NAMESPACE {
            continue;
        }
        let i = match child.l.as_str() {
            "dbPr" => 0,
            "olapPr" => 1,
            "webPr" => 2,
            "textPr" => 3,
            "parameters" => 4,
            "extLst" => 5,
            _ => continue,
        };
        if i < order {
            return Err(invalid("connection children out of order"));
        }
        order = i;
        match i {
            0 => set(&mut c.database, parse_db(child)?)?,
            1 => set(&mut c.olap, parse_olap(child)?)?,
            2 => set(&mut c.web, parse_web(child)?)?,
            3 => set(&mut c.text, parse_text(child)?)?,
            4 => set(&mut c.parameters, parse_parameters(child)?)?,
            5 => set(&mut c.extension_xml, node_xml(child, false)?)?,
            _ => return Err(invalid("unexpected connection child index")),
        }
    }
    Ok(c)
}
fn parse_db(n: &Node) -> Result<DatabaseProperties> {
    let v = DatabaseProperties {
        connection: req(n, "connection")?,
        command: aopt(n, "command")?,
        server_command: aopt(n, "serverCommand")?,
        command_type: u32opt(n, "commandType")?,
    };
    only(
        n,
        &["connection", "command", "serverCommand", "commandType"],
    )?;
    leaf(n)?;
    Ok(v)
}
fn parse_olap(n: &Node) -> Result<OlapProperties> {
    let v = OlapProperties {
        local: bopt(n, "local")?,
        local_connection: aopt(n, "localConnection")?,
        local_refresh: bopt(n, "localRefresh")?,
        send_locale: bopt(n, "sendLocale")?,
        row_drill_count: u32opt(n, "rowDrillCount")?,
        server_fill: bopt(n, "serverFill")?,
        server_number_format: bopt(n, "serverNumberFormat")?,
        server_font: bopt(n, "serverFont")?,
        server_font_color: bopt(n, "serverFontColor")?,
    };
    only(
        n,
        &[
            "local",
            "localConnection",
            "localRefresh",
            "sendLocale",
            "rowDrillCount",
            "serverFill",
            "serverNumberFormat",
            "serverFont",
            "serverFontColor",
        ],
    )?;
    leaf(n)?;
    Ok(v)
}
fn parse_web(n: &Node) -> Result<WebQueryProperties> {
    let mut v = WebQueryProperties {
        xml_source: bopt(n, "xml")?,
        source_data: bopt(n, "sourceData")?,
        parse_pre: bopt(n, "parsePre")?,
        consecutive: bopt(n, "consecutive")?,
        first_row: bopt(n, "firstRow")?,
        excel97: bopt(n, "xl97")?,
        text_dates: bopt(n, "textDates")?,
        excel2000: bopt(n, "xl2000")?,
        url: aopt(n, "url")?,
        post: aopt(n, "post")?,
        html_tables: bopt(n, "htmlTables")?,
        html_format: aopt(n, "htmlFormat")?.map(parse_html).transpose()?,
        edit_page: aopt(n, "editPage")?,
        tables: None,
    };
    only(
        n,
        &[
            "xml",
            "sourceData",
            "parsePre",
            "consecutive",
            "firstRow",
            "xl97",
            "textDates",
            "xl2000",
            "url",
            "post",
            "htmlTables",
            "htmlFormat",
            "editPage",
        ],
    )?;
    let c = kids(n)?;
    if c.len() > 1 {
        return Err(invalid("webPr permits one tables child"));
    }
    if let Some(t) = c.first() {
        expect(t, "tables")?;
        v.tables = Some(parse_tables(t)?);
    }
    Ok(v)
}
fn parse_tables(n: &Node) -> Result<Vec<WebTableSelector>> {
    let count = u32opt(n, "count")?;
    only(n, &["count"])?;
    let mut out = Vec::new();
    for c in kids(n)? {
        if out.len() >= MAX_WEB_TABLES {
            return Err(invalid("web table selector limit exceeded"));
        }
        expect_any(c)?;
        out.push(match c.l.as_str() {
            "m" => {
                noattrs(c)?;
                leaf(c)?;
                WebTableSelector::Missing
            },
            "s" => {
                let v = req(c, "v")?;
                only(c, &["v"])?;
                leaf(c)?;
                WebTableSelector::String(v)
            },
            "x" => {
                let v = u32req(c, "v")?;
                only(c, &["v"])?;
                leaf(c)?;
                WebTableSelector::Index(v)
            },
            _ => return Err(invalid("invalid web table selector")),
        });
    }
    if out.is_empty() {
        return Err(invalid("tables requires a selector"));
    }
    check_count(count, out.len(), "tables")?;
    Ok(out)
}
fn parse_text(n: &Node) -> Result<TextImportProperties> {
    let mut v = TextImportProperties {
        prompt: bopt(n, "prompt")?,
        file_type: aopt(n, "fileType")?.map(parse_file).transpose()?,
        code_page: u32opt(n, "codePage")?,
        character_set: aopt(n, "characterSet")?,
        first_row: u32opt(n, "firstRow")?,
        source_file: aopt(n, "sourceFile")?,
        delimited: bopt(n, "delimited")?,
        decimal: aopt(n, "decimal")?,
        thousands: aopt(n, "thousands")?,
        tab: bopt(n, "tab")?,
        space: bopt(n, "space")?,
        comma: bopt(n, "comma")?,
        semicolon: bopt(n, "semicolon")?,
        consecutive: bopt(n, "consecutive")?,
        qualifier: aopt(n, "qualifier")?.map(parse_qualifier).transpose()?,
        delimiter: aopt(n, "delimiter")?,
        fields: None,
    };
    only(
        n,
        &[
            "prompt",
            "fileType",
            "codePage",
            "characterSet",
            "firstRow",
            "sourceFile",
            "delimited",
            "decimal",
            "thousands",
            "tab",
            "space",
            "comma",
            "semicolon",
            "consecutive",
            "qualifier",
            "delimiter",
        ],
    )?;
    let c = kids(n)?;
    if c.len() > 1 {
        return Err(invalid("textPr permits one textFields child"));
    }
    if let Some(f) = c.first() {
        expect(f, "textFields")?;
        let count = u32opt(f, "count")?;
        only(f, &["count"])?;
        let mut fields = Vec::new();
        for e in kids(f)? {
            if fields.len() >= MAX_TEXT_FIELDS {
                return Err(invalid("text field limit exceeded"));
            }
            expect(e, "textField")?;
            fields.push(TextField {
                field_type: aopt(e, "type")?.map(parse_field).transpose()?,
                position: u32opt(e, "position")?,
            });
            only(e, &["type", "position"])?;
            leaf(e)?;
        }
        if fields.is_empty() {
            return Err(invalid("textFields requires a textField"));
        }
        check_count(count, fields.len(), "textFields")?;
        v.fields = Some(fields);
    }
    Ok(v)
}
fn parse_parameters(n: &Node) -> Result<Vec<ConnectionParameter>> {
    let count = u32opt(n, "count")?;
    only(n, &["count"])?;
    let mut out = Vec::new();
    for p in kids(n)? {
        if out.len() >= MAX_PARAMETERS {
            return Err(invalid("parameter limit exceeded"));
        }
        expect(p, "parameter")?;
        let double = match aopt(p, "double")? {
            Some(x) => {
                let v = x
                    .parse::<f64>()
                    .map_err(|_source| invalid("invalid parameter double"))?;
                if !v.is_finite() {
                    return Err(invalid("non-finite parameter double"));
                }
                Some(v)
            },
            None => None,
        };
        out.push(ConnectionParameter {
            name: aopt(p, "name")?,
            sql_type: i32opt(p, "sqlType")?,
            parameter_type: aopt(p, "parameterType")?
                .map(parse_parameter_type)
                .transpose()?,
            refresh_on_change: bopt(p, "refreshOnChange")?,
            prompt: aopt(p, "prompt")?,
            boolean: bopt(p, "boolean")?,
            double,
            integer: i32opt(p, "integer")?,
            string: aopt(p, "string")?,
            cell: aopt(p, "cell")?,
        });
        only(
            p,
            &[
                "name",
                "sqlType",
                "parameterType",
                "refreshOnChange",
                "prompt",
                "boolean",
                "double",
                "integer",
                "string",
                "cell",
            ],
        )?;
        leaf(p)?;
    }
    if out.is_empty() {
        return Err(invalid("parameters requires a parameter"));
    }
    check_count(count, out.len(), "parameters")?;
    Ok(out)
}

pub(super) struct BoundedXml {
    pub(super) bytes: Vec<u8>,
}

impl BoundedXml {
    pub(super) fn new() -> Self {
        Self { bytes: Vec::new() }
    }

    fn push_str(&mut self, value: &str) -> Result<()> {
        self.push_bytes(value.as_bytes())
    }

    fn push_char(&mut self, value: char) -> Result<()> {
        let mut encoded = [0; 4];
        let length = value.encode_utf8(&mut encoded).len();
        self.push_bytes(&encoded[..length])
    }

    pub(super) fn push_bytes(&mut self, value: &[u8]) -> Result<()> {
        let length = self
            .bytes
            .len()
            .checked_add(value.len())
            .ok_or_else(|| invalid("serialized connections length overflows"))?;
        if length > MAX_XML_BYTES {
            return Err(invalid("serialized connections part exceeds 16 MiB"));
        }
        self.bytes
            .try_reserve_exact(value.len())
            .map_err(|_source| invalid("serialized connections output allocation failed"))?;
        self.bytes.extend_from_slice(value);
        Ok(())
    }

    fn finish(self) -> Vec<u8> {
        self.bytes
    }
}

fn write_connection(x: &mut BoundedXml, c: &Connection, s: bool) -> Result<()> {
    x.push_str("<connection")?;
    num(x, "id", c.id)?;
    str_opt(x, "sourceFile", c.source_file.as_deref())?;
    str_opt(x, "odcFile", c.odc_file.as_deref())?;
    bool_opt(x, "keepAlive", c.keep_alive)?;
    num_opt(x, "interval", c.interval)?;
    str_opt(x, "name", c.name.as_deref())?;
    str_opt(x, "description", c.description.as_deref())?;
    num_opt(x, "type", c.connection_type)?;
    num_opt(x, "reconnectionMethod", c.reconnection_method)?;
    num(x, "refreshedVersion", c.refreshed_version)?;
    num_opt(x, "minRefreshableVersion", c.min_refreshable_version)?;
    bool_opt(x, "savePassword", c.save_password)?;
    bool_opt(x, "new", c.new_connection)?;
    bool_opt(x, "deleted", c.deleted)?;
    bool_opt(x, "onlyUseConnectionFile", c.only_use_connection_file)?;
    bool_opt(x, "background", c.background)?;
    bool_opt(x, "refreshOnLoad", c.refresh_on_load)?;
    bool_opt(x, "saveData", c.save_data)?;
    if let Some(v) = c.credentials {
        attr(x, "credentials", credentials_str(v))?;
    }
    str_opt(x, "singleSignOnId", c.single_sign_on_id.as_deref())?;
    if c.database.is_none()
        && c.olap.is_none()
        && c.web.is_none()
        && c.text.is_none()
        && c.parameters.is_none()
        && c.extension_xml.is_none()
    {
        x.push_str("/>")?;
        return Ok(());
    }
    x.push_char('>')?;
    if let Some(v) = &c.database {
        x.push_str("<dbPr")?;
        attr(x, "connection", &v.connection)?;
        str_opt(x, "command", v.command.as_deref())?;
        str_opt(x, "serverCommand", v.server_command.as_deref())?;
        num_opt(x, "commandType", v.command_type)?;
        x.push_str("/>")?;
    }
    if let Some(v) = &c.olap {
        write_olap(x, v)?;
    }
    if let Some(v) = &c.web {
        write_web(x, v)?;
    }
    if let Some(v) = &c.text {
        write_text(x, v)?;
    }
    if let Some(v) = &c.parameters {
        write_parameters(x, v)?;
    }
    if let Some(v) = &c.extension_xml {
        opaque(x, v, s)?;
    }
    x.push_str("</connection>")?;
    Ok(())
}
fn write_olap(x: &mut BoundedXml, v: &OlapProperties) -> Result<()> {
    x.push_str("<olapPr")?;
    for (n, b) in [
        ("local", v.local),
        ("localRefresh", v.local_refresh),
        ("sendLocale", v.send_locale),
        ("serverFill", v.server_fill),
        ("serverNumberFormat", v.server_number_format),
        ("serverFont", v.server_font),
        ("serverFontColor", v.server_font_color),
    ] {
        bool_opt(x, n, b)?;
    }
    str_opt(x, "localConnection", v.local_connection.as_deref())?;
    num_opt(x, "rowDrillCount", v.row_drill_count)?;
    x.push_str("/>")?;
    Ok(())
}
fn write_web(x: &mut BoundedXml, v: &WebQueryProperties) -> Result<()> {
    x.push_str("<webPr")?;
    for (n, b) in [
        ("xml", v.xml_source),
        ("sourceData", v.source_data),
        ("parsePre", v.parse_pre),
        ("consecutive", v.consecutive),
        ("firstRow", v.first_row),
        ("xl97", v.excel97),
        ("textDates", v.text_dates),
        ("xl2000", v.excel2000),
        ("htmlTables", v.html_tables),
    ] {
        bool_opt(x, n, b)?;
    }
    str_opt(x, "url", v.url.as_deref())?;
    str_opt(x, "post", v.post.as_deref())?;
    if let Some(h) = v.html_format {
        attr(x, "htmlFormat", html_str(h))?;
    }
    str_opt(x, "editPage", v.edit_page.as_deref())?;
    if let Some(t) = &v.tables {
        x.push_str("><tables")?;
        num(x, "count", t.len())?;
        x.push_char('>')?;
        for z in t {
            match z {
                WebTableSelector::Missing => x.push_str("<m/>")?,
                WebTableSelector::String(v) => {
                    x.push_str("<s")?;
                    attr(x, "v", v)?;
                    x.push_str("/>")?;
                },
                WebTableSelector::Index(v) => {
                    x.push_str("<x")?;
                    num(x, "v", *v)?;
                    x.push_str("/>")?;
                },
            }
        }
        x.push_str("</tables></webPr>")?;
    } else {
        x.push_str("/>")?;
    }
    Ok(())
}
fn write_text(x: &mut BoundedXml, v: &TextImportProperties) -> Result<()> {
    x.push_str("<textPr")?;
    bool_opt(x, "prompt", v.prompt)?;
    if let Some(z) = v.file_type {
        attr(x, "fileType", file_str(z))?;
    }
    num_opt(x, "codePage", v.code_page)?;
    str_opt(x, "characterSet", v.character_set.as_deref())?;
    num_opt(x, "firstRow", v.first_row)?;
    str_opt(x, "sourceFile", v.source_file.as_deref())?;
    for (n, b) in [
        ("delimited", v.delimited),
        ("tab", v.tab),
        ("space", v.space),
        ("comma", v.comma),
        ("semicolon", v.semicolon),
        ("consecutive", v.consecutive),
    ] {
        bool_opt(x, n, b)?;
    }
    str_opt(x, "decimal", v.decimal.as_deref())?;
    str_opt(x, "thousands", v.thousands.as_deref())?;
    if let Some(z) = v.qualifier {
        attr(x, "qualifier", qualifier_str(z))?;
    }
    str_opt(x, "delimiter", v.delimiter.as_deref())?;
    if let Some(f) = &v.fields {
        x.push_str("><textFields")?;
        num(x, "count", f.len())?;
        x.push_char('>')?;
        for z in f {
            x.push_str("<textField")?;
            if let Some(t) = z.field_type {
                attr(x, "type", field_str(t))?;
            }
            num_opt(x, "position", z.position)?;
            x.push_str("/>")?;
        }
        x.push_str("</textFields></textPr>")?;
    } else {
        x.push_str("/>")?;
    }
    Ok(())
}
fn write_parameters(x: &mut BoundedXml, v: &[ConnectionParameter]) -> Result<()> {
    x.push_str("<parameters")?;
    num(x, "count", v.len())?;
    x.push_char('>')?;
    for p in v {
        x.push_str("<parameter")?;
        str_opt(x, "name", p.name.as_deref())?;
        num_opt(x, "sqlType", p.sql_type)?;
        if let Some(z) = p.parameter_type {
            attr(x, "parameterType", parameter_str(z))?;
        }
        bool_opt(x, "refreshOnChange", p.refresh_on_change)?;
        str_opt(x, "prompt", p.prompt.as_deref())?;
        bool_opt(x, "boolean", p.boolean)?;
        if let Some(z) = p.double {
            let value = z.to_string();
            attr(x, "double", &value)?;
        }
        num_opt(x, "integer", p.integer)?;
        str_opt(x, "string", p.string.as_deref())?;
        str_opt(x, "cell", p.cell.as_deref())?;
        x.push_str("/>")?;
    }
    x.push_str("</parameters>")?;
    Ok(())
}

fn opaque(x: &mut BoundedXml, b: &[u8], strict: bool) -> Result<()> {
    parse_dom(b)?;
    let source = rewrite_namespace_declarations(b, strict)?;
    x.push_bytes(&source)?;
    Ok(())
}

fn rewrite_namespace_declarations(source: &[u8], strict: bool) -> Result<Vec<u8>> {
    let tree = SourceTree::parse_raw(source)?;
    let (from, to) = if strict {
        (CORE_NAMESPACE, STRICT_NAMESPACE)
    } else {
        (STRICT_NAMESPACE, CORE_NAMESPACE)
    };
    let mut edits = Vec::new();
    for node in &tree.nodes {
        for attribute in &node.attrs {
            if attribute.namespace_declaration
                && decode_source_attribute(source, attribute)?.as_str() == from
            {
                edits.push(SourceEdit {
                    range: attribute.value_start..attribute.value_end,
                    replacement: to.as_bytes().to_vec(),
                });
            }
        }
    }
    apply_source_edits(source, edits)
}
fn node_xml(n: &Node, s: bool) -> Result<Vec<u8>> {
    let mut x = String::new();
    node_write(&mut x, n, s)?;
    Ok(x.into_bytes())
}
fn node_write(x: &mut String, n: &Node, s: bool) -> Result<()> {
    x.push('<');
    x.push_str(&n.q);
    for (p, u) in n.context.effective_bindings()? {
        if p.is_empty() {
            x.push_str(" xmlns=\"");
        } else {
            x.push_str(" xmlns:");
            x.push_str(p);
            x.push_str("=\"");
        }
        esc(
            x,
            if s && u == CORE_NAMESPACE {
                STRICT_NAMESPACE
            } else if !s && u == STRICT_NAMESPACE {
                CORE_NAMESPACE
            } else {
                u
            },
        );
        x.push('"');
    }
    for a in &n.attrs {
        x.push(' ');
        x.push_str(&a.q);
        x.push_str("=\"");
        esc(x, &a.v);
        x.push('"');
    }
    if n.content.is_empty() {
        x.push_str("/>");
        return Ok(());
    }
    x.push('>');
    for c in &n.content {
        match c {
            Content::Node(n) => node_write(x, n, s)?,
            Content::Text(v) => text_escape(x, v),
            Content::CData(v) => {
                x.push_str("<![CDATA[");
                x.push_str(v);
                x.push_str("]]>");
            },
            Content::Comment(v) => {
                x.push_str("<!--");
                x.push_str(v);
                x.push_str("-->");
            },
        }
    }
    x.push_str("</");
    x.push_str(&n.q);
    x.push('>');
    Ok(())
}

pub(super) fn kids(n: &Node) -> Result<Vec<&Node>> {
    let mut v = Vec::new();
    for c in &n.content {
        match c {
            Content::Node(x) => v.push(x),
            Content::Text(x) if x.trim().is_empty() => {},
            Content::Comment(_) => {},
            Content::Text(_) | Content::CData(_) => {
                return Err(invalid("unexpected text in typed connections"));
            },
        }
    }
    Ok(v)
}
fn leaf(n: &Node) -> Result<()> {
    if kids(n)?.is_empty() {
        Ok(())
    } else {
        Err(invalid("connection leaf has children"))
    }
}
pub(super) fn expect(n: &Node, l: &str) -> Result<()> {
    if (n.ns.as_ref() == CORE_NAMESPACE || n.ns.as_ref() == STRICT_NAMESPACE) && n.l == l {
        Ok(())
    } else {
        Err(invalid(format!("expected SpreadsheetML {l}")))
    }
}
fn expect_any(n: &Node) -> Result<()> {
    if n.ns.as_ref() == CORE_NAMESPACE || n.ns.as_ref() == STRICT_NAMESPACE {
        Ok(())
    } else {
        Err(invalid("expected SpreadsheetML child"))
    }
}
fn aopt(n: &Node, l: &str) -> Result<Option<String>> {
    let mut v = None;
    for a in &n.attrs {
        if a.ns.is_empty() && a.l == l {
            if v.is_some() {
                return Err(invalid("duplicate attribute"));
            }
            bounded(&a.v)?;
            v = Some(a.v.clone());
        }
    }
    Ok(v)
}
pub(super) fn req(n: &Node, l: &str) -> Result<String> {
    aopt(n, l)?.ok_or_else(|| invalid(format!("missing required attribute '{l}'")))
}
fn bopt(n: &Node, l: &str) -> Result<Option<bool>> {
    match aopt(n, l)?.as_deref() {
        None => Ok(None),
        Some("1" | "true") => Ok(Some(true)),
        Some("0" | "false") => Ok(Some(false)),
        _ => Err(invalid(format!("invalid boolean '{l}'"))),
    }
}
fn u32opt(n: &Node, l: &str) -> Result<Option<u32>> {
    aopt(n, l)?
        .map(|x| {
            x.parse()
                .map_err(|_source| invalid(format!("invalid u32 '{l}'")))
        })
        .transpose()
}
pub(super) fn u32req(n: &Node, l: &str) -> Result<u32> {
    u32opt(n, l)?.ok_or_else(|| invalid(format!("missing u32 '{l}'")))
}
fn u8opt(n: &Node, l: &str) -> Result<Option<u8>> {
    aopt(n, l)?
        .map(|x| {
            x.parse()
                .map_err(|_source| invalid(format!("invalid u8 '{l}'")))
        })
        .transpose()
}
fn u8req(n: &Node, l: &str) -> Result<u8> {
    u8opt(n, l)?.ok_or_else(|| invalid(format!("missing u8 '{l}'")))
}
fn i32opt(n: &Node, l: &str) -> Result<Option<i32>> {
    aopt(n, l)?
        .map(|x| {
            x.parse()
                .map_err(|_source| invalid(format!("invalid i32 '{l}'")))
        })
        .transpose()
}
fn only(n: &Node, a: &[&str]) -> Result<()> {
    // Attribute extensions are deliberately opaque. The package transaction
    // retains their original source span while typed access continues to
    // validate every known attribute.
    let _ = (n, a);
    Ok(())
}

pub(super) fn only_unqualified(n: &Node, allowed: &[&str]) -> Result<()> {
    for attribute in &n.attrs {
        if attribute.ns.is_empty() && !allowed.contains(&attribute.l.as_str()) {
            return Err(invalid(format!("unexpected attribute '{}'", attribute.q)));
        }
    }
    Ok(())
}
fn noattrs(n: &Node) -> Result<()> {
    only(n, &[])
}
fn set<T>(s: &mut Option<T>, v: T) -> Result<()> {
    if s.replace(v).is_some() {
        Err(invalid("duplicate connection property"))
    } else {
        Ok(())
    }
}
fn check_count(c: Option<u32>, actual: usize, n: &str) -> Result<()> {
    if c.is_some_and(|x| x as usize != actual) {
        Err(invalid(format!("{n} count mismatch")))
    } else {
        Ok(())
    }
}
fn split(q: &str) -> Result<(&str, &str)> {
    if let Some((p, l)) = q.split_once(':') {
        if l.is_empty() || l.contains(':') {
            return Err(invalid("invalid QName"));
        }
        Ok((p, l))
    } else {
        Ok(("", q))
    }
}
fn resolve(context: &NamespaceContext, p: &str) -> Result<NamespaceUri> {
    context
        .resolve_uri(p)
        .cloned()
        .ok_or_else(|| invalid(format!("unbound prefix '{p}'")))
}
pub(super) fn bounded(v: &str) -> Result<()> {
    if v.len() > MAX_STRING_BYTES {
        Err(invalid("connection string exceeds 1 MiB"))
    } else {
        Ok(())
    }
}
fn attr(x: &mut BoundedXml, n: &str, v: &str) -> Result<()> {
    x.push_char(' ')?;
    x.push_str(n)?;
    x.push_str("=\"")?;
    esc_bounded(x, v)?;
    x.push_char('"')
}
fn str_opt(x: &mut BoundedXml, n: &str, v: Option<&str>) -> Result<()> {
    if let Some(v) = v {
        attr(x, n, v)?;
    }
    Ok(())
}
fn bool_opt(x: &mut BoundedXml, n: &str, v: Option<bool>) -> Result<()> {
    if let Some(v) = v {
        attr(x, n, if v { "1" } else { "0" })?;
    }
    Ok(())
}
fn num<T: std::fmt::Display>(x: &mut BoundedXml, n: &str, v: T) -> Result<()> {
    attr(x, n, &v.to_string())
}
fn num_opt<T: std::fmt::Display>(x: &mut BoundedXml, n: &str, v: Option<T>) -> Result<()> {
    if let Some(v) = v {
        num(x, n, v)?;
    }
    Ok(())
}
fn esc_bounded(x: &mut BoundedXml, v: &str) -> Result<()> {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;")?,
            '<' => x.push_str("&lt;")?,
            '"' => x.push_str("&quot;")?,
            '\r' => x.push_str("&#xD;")?,
            '\n' => x.push_str("&#xA;")?,
            '\t' => x.push_str("&#x9;")?,
            _ => x.push_char(c)?,
        }
    }
    Ok(())
}
fn esc(x: &mut String, v: &str) {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;"),
            '<' => x.push_str("&lt;"),
            '"' => x.push_str("&quot;"),
            '\r' => x.push_str("&#xD;"),
            '\n' => x.push_str("&#xA;"),
            '\t' => x.push_str("&#x9;"),
            _ => x.push(c),
        }
    }
}
fn text_escape(x: &mut String, v: &str) {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;"),
            '<' => x.push_str("&lt;"),
            '>' => x.push_str("&gt;"),
            _ => x.push(c),
        }
    }
}
macro_rules! en{($p:ident,$w:ident,$t:ty,$($s:literal=>$v:path),+)=>{fn $p(s:String)->Result<$t>{match s.as_str(){$($s=>Ok($v),)+_=>Err(invalid(format!("invalid enumeration '{s}'")))}}fn $w(v:$t)->&'static str{match v{$($v=>$s,)+}}}}
en!(parse_credentials,credentials_str,CredentialsMethod,"integrated"=>CredentialsMethod::Integrated,"none"=>CredentialsMethod::None,"stored"=>CredentialsMethod::Stored);
en!(parse_html,html_str,HtmlFormatting,"none"=>HtmlFormatting::None,"rtf"=>HtmlFormatting::RichText,"all"=>HtmlFormatting::All);
en!(parse_file,file_str,TextFileType,"mac"=>TextFileType::Mac,"win"=>TextFileType::Windows,"dos"=>TextFileType::Dos);
en!(parse_qualifier,qualifier_str,TextQualifier,"doubleQuote"=>TextQualifier::DoubleQuote,"singleQuote"=>TextQualifier::SingleQuote,"none"=>TextQualifier::None);
en!(parse_parameter_type,parameter_str,ParameterType,"prompt"=>ParameterType::Prompt,"value"=>ParameterType::Value,"cell"=>ParameterType::Cell);
en!(parse_field,field_str,TextFieldType,"general"=>TextFieldType::General,"text"=>TextFieldType::Text,"MDY"=>TextFieldType::MonthDayYear,"DMY"=>TextFieldType::DayMonthYear,"YMD"=>TextFieldType::YearMonthDay,"MYD"=>TextFieldType::MonthYearDay,"DYM"=>TextFieldType::DayYearMonth,"YDM"=>TextFieldType::YearDayMonth,"skip"=>TextFieldType::Skip,"EMD"=>TextFieldType::EastAsianYearMonthDay);

fn xml_error(e: impl std::fmt::Display) -> Box<dyn std::error::Error + Send + Sync> {
    invalid(e.to_string())
}

/// Patch typed connection fields inside their original XML spans.
///
/// Connection and query-table parts contain many producer extensions. The
/// transaction therefore changes scalar attributes in place and replaces only
/// the known property child whose typed value changed. Unrelated source bytes
/// remain untouched; structural collection edits retain existing connection
/// blocks and append new canonical blocks. An `extension_xml` change is an
/// explicit whole-`extLst` owner replacement or removal: replacement XML owns
/// its complete attribute and child set, so old opaque markup is not merged.
pub(super) fn patch_connections_source(
    source: &[u8],
    before: &Connections,
    after: &Connections,
    strict: bool,
) -> Result<Vec<u8>> {
    if before == after {
        return Ok(source.to_vec());
    }
    let tree = SourceTree::parse(source)?;
    if tree.nodes[tree.root].local != "connections"
        || !matches!(
            &*tree.nodes[tree.root].namespace,
            CORE_NAMESPACE | STRICT_NAMESPACE
        )
        || tree.nodes[tree.root].self_closing
    {
        return Err(invalid(
            "connection source has no editable SpreadsheetML connections root",
        ));
    }
    let nodes = tree.connection_nodes()?;
    if nodes.len() != before.connections.len() {
        return Err(invalid(
            "connection source has an ambiguous active connection layout",
        ));
    }
    let mut source_ids = Vec::new();
    source_ids
        .try_reserve_exact(nodes.len())
        .map_err(|_| invalid("connection source ID index allocation failed"))?;
    let mut source_by_id = HashMap::new();
    source_by_id
        .try_reserve(nodes.len())
        .map_err(|_| invalid("connection source ID index allocation failed"))?;
    for (index, node) in nodes.iter().enumerate() {
        let id = tree
            .attribute(source, *node, "id")?
            .ok_or_else(|| invalid("connection source is missing its id"))?
            .parse::<u32>()
            .map_err(|_source| invalid("connection source has an invalid id"))?;
        if source_by_id.insert(id, index).is_some() {
            return Err(invalid("duplicate connection id"));
        }
        source_ids.push(id);
    }
    let mut before_by_id = HashMap::new();
    before_by_id
        .try_reserve(before.connections.len())
        .map_err(|_| invalid("connection before ID index allocation failed"))?;
    for (index, connection) in before.connections.iter().enumerate() {
        if before_by_id.insert(connection.id, index).is_some() {
            return Err(invalid("duplicate connection id"));
        }
    }
    let mut after_by_id = HashMap::new();
    after_by_id
        .try_reserve(after.connections.len())
        .map_err(|_| invalid("connection after ID index allocation failed"))?;
    for (index, connection) in after.connections.iter().enumerate() {
        if after_by_id.insert(connection.id, index).is_some() {
            return Err(invalid("duplicate connection id"));
        }
    }
    let mut edits = Vec::new();
    let mut extension_replacements = Vec::new();
    let insertion = nodes.last().and_then(|node| {
        tree.nodes[*node]
            .parent
            .map(|parent| (parent, tree.nodes[*node].end))
    });
    for connection in &after.connections {
        if let Some(&before_index) = before_by_id.get(&connection.id) {
            let source_node = *source_by_id
                .get(&connection.id)
                .ok_or_else(|| invalid("connection source identity changed"))?;
            let node = nodes[source_node];
            patch_connection_source(
                &tree,
                source,
                node,
                &before.connections[before_index],
                connection,
                strict,
                &mut edits,
                &mut extension_replacements,
            )?;
        } else {
            let (parent, offset) = insertion.ok_or_else(|| {
                invalid("connection source has no safe connection insertion point")
            })?;
            let canonical = canonical_connection(connection, strict)?;
            let replacement = qualify_fragment(canonical.clone(), &tree, parent)?;
            if connection.extension_xml.is_some() {
                let extension = canonical_extension(&canonical, strict)?
                    .ok_or_else(|| invalid("canonical connection lost its extLst"))?;
                record_extension_publication(
                    &mut extension_replacements,
                    connection.id,
                    &extension,
                )?;
            }
            edits.push(SourceEdit {
                range: offset..offset,
                replacement,
            });
        }
    }
    for (index, node) in nodes.iter().enumerate() {
        if !after_by_id.contains_key(&source_ids[index]) {
            edits.push(SourceEdit {
                range: tree.nodes[*node].start..tree.nodes[*node].end,
                replacement: Vec::new(),
            });
        }
    }
    let updated = apply_source_edits(source, edits)?;
    let updated = reorder_connections_source(&updated, &after.connections)?;
    verify_extension_publications(&updated, &extension_replacements)?;
    Ok(updated)
}

/// Project staged opaque extension owners from the already reopened candidate.
/// Source publication is independently checked by `patch_connections_source`
/// before this projection is used, so this only accounts for contextual MCE,
/// namespace, and processing-instruction normalization.
pub(super) fn normalize_connections_source_projection(
    actual: &Connections,
    value: &mut Connections,
) -> Result<()> {
    if !value
        .connections
        .iter()
        .any(|connection| connection.extension_xml.is_some())
    {
        return Ok(());
    }
    let mut by_id = HashMap::new();
    by_id
        .try_reserve(actual.connections.len())
        .map_err(|_| invalid("connection projection index allocation failed"))?;
    for connection in &actual.connections {
        if by_id.insert(connection.id, connection).is_some() {
            return Err(invalid("connection projection has duplicate IDs"));
        }
    }
    for connection in &mut value.connections {
        if connection.extension_xml.is_none() {
            continue;
        }
        let reopened = by_id
            .get(&connection.id)
            .ok_or_else(|| invalid("connection projection is missing a staged ID"))?;
        if connection.extension_xml.as_deref() != reopened.extension_xml.as_deref() {
            connection.extension_xml = reopened.extension_xml.clone();
        }
    }
    Ok(())
}

#[derive(Debug)]
struct PublishedExtension {
    id: u32,
    bytes: Vec<u8>,
}

fn record_extension_publication(
    publications: &mut Vec<PublishedExtension>,
    id: u32,
    bytes: &[u8],
) -> Result<()> {
    publications
        .try_reserve(1)
        .map_err(|_| invalid("connection extension publication allocation failed"))?;
    let mut copy = Vec::new();
    copy.try_reserve_exact(bytes.len())
        .map_err(|_| invalid("connection extension publication allocation failed"))?;
    copy.extend_from_slice(bytes);
    publications.push(PublishedExtension { id, bytes: copy });
    Ok(())
}

fn verify_extension_publications(source: &[u8], publications: &[PublishedExtension]) -> Result<()> {
    if publications.is_empty() {
        return Ok(());
    }
    let tree = SourceTree::parse(source)?;
    let nodes = tree.connection_nodes()?;
    let mut by_id = HashMap::new();
    by_id
        .try_reserve(nodes.len())
        .map_err(|_| invalid("connection publication index allocation failed"))?;
    for node in nodes {
        let id = tree
            .attribute(source, node, "id")?
            .ok_or_else(|| invalid("connection source is missing its id"))?
            .parse::<u32>()
            .map_err(|_| invalid("connection source has an invalid id"))?;
        if by_id.insert(id, node).is_some() {
            return Err(invalid("connection source has duplicate IDs"));
        }
    }
    for publication in publications {
        let node = *by_id
            .get(&publication.id)
            .ok_or_else(|| invalid("published connection ID is missing"))?;
        let extension = tree
            .child(node, "extLst")?
            .ok_or_else(|| invalid("published connection is missing its extLst"))?;
        let actual = source
            .get(tree.nodes[extension].start..tree.nodes[extension].end)
            .ok_or_else(|| invalid("published extLst span is invalid"))?;
        if actual != publication.bytes.as_slice() {
            return Err(invalid("published extLst differs from its requested owner"));
        }
    }
    Ok(())
}

fn reorder_connections_source(source: &[u8], desired: &[Connection]) -> Result<Vec<u8>> {
    let tree = SourceTree::parse(source)?;
    let nodes = tree.connection_nodes()?;
    if nodes.len() != desired.len() {
        return Err(invalid(
            "connection source has an ambiguous active connection layout",
        ));
    }
    let mut source_ids = Vec::new();
    source_ids
        .try_reserve_exact(nodes.len())
        .map_err(|_source| invalid("connection reorder allocation failed"))?;
    let mut source_by_id = HashMap::new();
    source_by_id
        .try_reserve(nodes.len())
        .map_err(|_source| invalid("connection reorder ID index allocation failed"))?;
    for (index, node) in nodes.iter().enumerate() {
        let id = tree
            .attribute(source, *node, "id")?
            .ok_or_else(|| invalid("connection source is missing its id"))?
            .parse::<u32>()
            .map_err(|_source| invalid("connection source has an invalid id"))?;
        if source_by_id.insert(id, index).is_some() {
            return Err(invalid("duplicate connection id"));
        }
        source_ids.push(id);
    }
    let mut desired_ids = HashSet::new();
    desired_ids
        .try_reserve(desired.len())
        .map_err(|_source| invalid("connection reorder ID set allocation failed"))?;
    if desired
        .iter()
        .any(|connection| !desired_ids.insert(connection.id))
    {
        return Err(invalid("connection reorder is not a source permutation"));
    }
    if source_ids
        .iter()
        .zip(desired)
        .all(|(source_id, connection)| *source_id == connection.id)
    {
        return apply_source_edits(source, Vec::new());
    }
    let parent = nodes
        .first()
        .and_then(|node| tree.nodes[*node].parent)
        .ok_or_else(|| invalid("connection source has no reorder parent"))?;
    // Moving a connection across direct/MCE parents would change the source
    // branch selected by markup-compatibility processing, so preserve the
    // source topology and refuse that ambiguous reorder.
    if nodes
        .iter()
        .any(|node| tree.nodes[*node].parent != Some(parent))
    {
        return Err(invalid(
            "connection reorder crosses source structural branches",
        ));
    }
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(nodes.len())
        .map_err(|_source| invalid("connection reorder allocation failed"))?;
    for (position, connection) in desired.iter().enumerate() {
        let source_node = *source_by_id
            .get(&connection.id)
            .ok_or_else(|| invalid("connection reorder is not a source permutation"))?;
        if source_node == position {
            continue;
        }
        let node = nodes[source_node];
        edits.push(SourceEdit {
            range: tree.nodes[nodes[position]].start..tree.nodes[nodes[position]].end,
            replacement: source
                .get(tree.nodes[node].start..tree.nodes[node].end)
                .ok_or_else(|| invalid("connection source reorder span is invalid"))?
                .to_vec(),
        });
    }
    apply_source_edits(source, edits)
}

fn patch_connection_source(
    tree: &SourceTree,
    source: &[u8],
    node: usize,
    before: &Connection,
    after: &Connection,
    strict: bool,
    edits: &mut Vec<SourceEdit>,
    extension_replacements: &mut Vec<PublishedExtension>,
) -> Result<()> {
    if tree.nodes[node].local != "connection"
        || !matches!(
            &*tree.nodes[node].namespace,
            CORE_NAMESPACE | STRICT_NAMESPACE
        )
    {
        return Err(invalid("connection source has an invalid root"));
    }
    patch_optional_attr(
        tree,
        node,
        "id",
        Some(&before.id.to_string()),
        Some(&after.id.to_string()),
        edits,
    )?;
    patch_optional_attr(
        tree,
        node,
        "sourceFile",
        before.source_file.as_deref(),
        after.source_file.as_deref(),
        edits,
    )?;
    patch_optional_attr(
        tree,
        node,
        "odcFile",
        before.odc_file.as_deref(),
        after.odc_file.as_deref(),
        edits,
    )?;
    patch_bool(
        tree,
        node,
        "keepAlive",
        before.keep_alive,
        after.keep_alive,
        edits,
    )?;
    patch_number(
        tree,
        node,
        "interval",
        before.interval,
        after.interval,
        edits,
    )?;
    patch_optional_attr(
        tree,
        node,
        "name",
        before.name.as_deref(),
        after.name.as_deref(),
        edits,
    )?;
    patch_optional_attr(
        tree,
        node,
        "description",
        before.description.as_deref(),
        after.description.as_deref(),
        edits,
    )?;
    patch_number(
        tree,
        node,
        "type",
        before.connection_type,
        after.connection_type,
        edits,
    )?;
    patch_number(
        tree,
        node,
        "reconnectionMethod",
        before.reconnection_method,
        after.reconnection_method,
        edits,
    )?;
    patch_optional_attr(
        tree,
        node,
        "refreshedVersion",
        Some(&before.refreshed_version.to_string()),
        Some(&after.refreshed_version.to_string()),
        edits,
    )?;
    patch_number(
        tree,
        node,
        "minRefreshableVersion",
        before.min_refreshable_version,
        after.min_refreshable_version,
        edits,
    )?;
    patch_bool(
        tree,
        node,
        "savePassword",
        before.save_password,
        after.save_password,
        edits,
    )?;
    patch_bool(
        tree,
        node,
        "new",
        before.new_connection,
        after.new_connection,
        edits,
    )?;
    patch_bool(tree, node, "deleted", before.deleted, after.deleted, edits)?;
    patch_bool(
        tree,
        node,
        "onlyUseConnectionFile",
        before.only_use_connection_file,
        after.only_use_connection_file,
        edits,
    )?;
    patch_bool(
        tree,
        node,
        "background",
        before.background,
        after.background,
        edits,
    )?;
    patch_bool(
        tree,
        node,
        "refreshOnLoad",
        before.refresh_on_load,
        after.refresh_on_load,
        edits,
    )?;
    patch_bool(
        tree,
        node,
        "saveData",
        before.save_data,
        after.save_data,
        edits,
    )?;
    patch_optional_attr(
        tree,
        node,
        "credentials",
        before.credentials.map(credentials_str),
        after.credentials.map(credentials_str),
        edits,
    )?;
    patch_optional_attr(
        tree,
        node,
        "singleSignOnId",
        before.single_sign_on_id.as_deref(),
        after.single_sign_on_id.as_deref(),
        edits,
    )?;

    let mut self_closing_children = Vec::new();
    patch_child(
        tree,
        source,
        node,
        "dbPr",
        before,
        after,
        before.database != after.database,
        after
            .database
            .is_some()
            .then(|| canonical_child(after, "dbPr", strict))
            .transpose()?,
        edits,
        &mut self_closing_children,
        extension_replacements,
    )?;
    patch_child(
        tree,
        source,
        node,
        "olapPr",
        before,
        after,
        before.olap != after.olap,
        after
            .olap
            .is_some()
            .then(|| canonical_child(after, "olapPr", strict))
            .transpose()?,
        edits,
        &mut self_closing_children,
        extension_replacements,
    )?;
    patch_child(
        tree,
        source,
        node,
        "webPr",
        before,
        after,
        before.web != after.web,
        after
            .web
            .is_some()
            .then(|| canonical_child(after, "webPr", strict))
            .transpose()?,
        edits,
        &mut self_closing_children,
        extension_replacements,
    )?;
    patch_child(
        tree,
        source,
        node,
        "textPr",
        before,
        after,
        before.text != after.text,
        after
            .text
            .is_some()
            .then(|| canonical_child(after, "textPr", strict))
            .transpose()?,
        edits,
        &mut self_closing_children,
        extension_replacements,
    )?;
    patch_child(
        tree,
        source,
        node,
        "parameters",
        before,
        after,
        before.parameters != after.parameters,
        after
            .parameters
            .is_some()
            .then(|| canonical_child(after, "parameters", strict))
            .transpose()?,
        edits,
        &mut self_closing_children,
        extension_replacements,
    )?;
    patch_child(
        tree,
        source,
        node,
        "extLst",
        before,
        after,
        before.extension_xml != after.extension_xml,
        after
            .extension_xml
            .as_ref()
            .map(|_| canonical_child(after, "extLst", strict))
            .transpose()?,
        edits,
        &mut self_closing_children,
        extension_replacements,
    )?;
    expand_self_closing(tree, source, node, self_closing_children, edits)?;
    Ok(())
}

fn canonical_connection(value: &Connection, strict: bool) -> Result<Vec<u8>> {
    let mut output = BoundedXml::new();
    write_connection(&mut output, value, strict)?;
    Ok(output.finish())
}

fn canonical_child(value: &Connection, name: &str, strict: bool) -> Result<Vec<u8>> {
    let source = canonical_connection(value, strict)?;
    let namespace = if strict {
        STRICT_NAMESPACE
    } else {
        CORE_NAMESPACE
    };
    let inherited = [(String::new(), namespace.to_owned())];
    let tree = SourceTree::parse_with_bindings(&source, &inherited)?;
    let node = tree
        .child(tree.root, name)?
        .ok_or_else(|| invalid(format!("canonical connection is missing '{name}'")))?;
    Ok(source[tree.nodes[node].start..tree.nodes[node].end].to_vec())
}

fn canonical_extension(source: &[u8], strict: bool) -> Result<Option<Vec<u8>>> {
    let namespace = if strict {
        STRICT_NAMESPACE
    } else {
        CORE_NAMESPACE
    };
    let inherited = [(String::new(), namespace.to_owned())];
    let tree = SourceTree::parse_with_bindings(source, &inherited)?;
    tree.child(tree.root, "extLst")?
        .map(|node| {
            source
                .get(tree.nodes[node].start..tree.nodes[node].end)
                .ok_or_else(|| invalid("canonical extLst span is invalid"))
                .map(ToOwned::to_owned)
        })
        .transpose()
}

fn patch_child(
    tree: &SourceTree,
    source: &[u8],
    parent: usize,
    name: &str,
    before: &Connection,
    after: &Connection,
    changed: bool,
    replacement: Option<Vec<u8>>,
    edits: &mut Vec<SourceEdit>,
    self_closing_children: &mut Vec<(String, Vec<u8>)>,
    extension_replacements: &mut Vec<PublishedExtension>,
) -> Result<()> {
    if !changed {
        return Ok(());
    }
    let existing = tree.child(parent, name)?;
    match (existing, replacement) {
        (Some(node), Some(replacement)) => {
            if patch_existing_child_source(tree, source, node, name, before, after, edits)? {
                return Ok(());
            }
            if name == "extLst" {
                let insertion_parent = tree.nodes[node].parent.unwrap_or(parent);
                let replacement = qualify_fragment(replacement, tree, insertion_parent)?;
                record_extension_publication(extension_replacements, after.id, &replacement)?;
                edits.push(SourceEdit {
                    range: tree.nodes[node].start..tree.nodes[node].end,
                    replacement,
                });
                return Ok(());
            }
            return Err(invalid(format!(
                "connection source cannot preserve structural '{name}' edits"
            )));
        },
        (Some(node), None) => edits.push(SourceEdit {
            range: tree.nodes[node].start..tree.nodes[node].end,
            replacement: Vec::new(),
        }),
        (None, Some(replacement)) => {
            if tree.nodes[parent].self_closing {
                let replacement = if name == "extLst" {
                    let replacement = qualify_fragment(replacement, tree, parent)?;
                    record_extension_publication(extension_replacements, after.id, &replacement)?;
                    replacement
                } else {
                    replacement
                };
                self_closing_children.push((name.to_owned(), replacement));
                return Ok(());
            }
            let insertion_parent = tree
                .semantic_insertion_parent(parent, name)
                .unwrap_or(parent);
            let offset = child_insertion_offset(
                tree,
                insertion_parent,
                tree.nodes[parent].namespace.as_ref(),
                name,
            );
            let replacement = qualify_fragment(replacement, tree, insertion_parent)?;
            if name == "extLst" {
                record_extension_publication(extension_replacements, after.id, &replacement)?;
            }
            edits.push(SourceEdit {
                range: offset..offset,
                replacement,
            });
        },
        (None, None) => {},
    }
    Ok(())
}

fn patch_existing_child_source(
    tree: &SourceTree,
    source: &[u8],
    node: usize,
    name: &str,
    before: &Connection,
    after: &Connection,
    edits: &mut Vec<SourceEdit>,
) -> Result<bool> {
    match name {
        "dbPr" => {
            let (Some(before), Some(after)) = (&before.database, &after.database) else {
                return Ok(false);
            };
            patch_optional_attr(
                tree,
                node,
                "connection",
                Some(&before.connection),
                Some(&after.connection),
                edits,
            )?;
            patch_optional_attr(
                tree,
                node,
                "command",
                before.command.as_deref(),
                after.command.as_deref(),
                edits,
            )?;
            patch_optional_attr(
                tree,
                node,
                "serverCommand",
                before.server_command.as_deref(),
                after.server_command.as_deref(),
                edits,
            )?;
            patch_number(
                tree,
                node,
                "commandType",
                before.command_type,
                after.command_type,
                edits,
            )?;
            Ok(true)
        },
        "olapPr" => {
            let (Some(before), Some(after)) = (&before.olap, &after.olap) else {
                return Ok(false);
            };
            patch_bool(tree, node, "local", before.local, after.local, edits)?;
            patch_optional_attr(
                tree,
                node,
                "localConnection",
                before.local_connection.as_deref(),
                after.local_connection.as_deref(),
                edits,
            )?;
            patch_bool(
                tree,
                node,
                "localRefresh",
                before.local_refresh,
                after.local_refresh,
                edits,
            )?;
            patch_bool(
                tree,
                node,
                "sendLocale",
                before.send_locale,
                after.send_locale,
                edits,
            )?;
            patch_number(
                tree,
                node,
                "rowDrillCount",
                before.row_drill_count,
                after.row_drill_count,
                edits,
            )?;
            patch_bool(
                tree,
                node,
                "serverFill",
                before.server_fill,
                after.server_fill,
                edits,
            )?;
            patch_bool(
                tree,
                node,
                "serverNumberFormat",
                before.server_number_format,
                after.server_number_format,
                edits,
            )?;
            patch_bool(
                tree,
                node,
                "serverFont",
                before.server_font,
                after.server_font,
                edits,
            )?;
            patch_bool(
                tree,
                node,
                "serverFontColor",
                before.server_font_color,
                after.server_font_color,
                edits,
            )?;
            Ok(true)
        },
        "webPr" => {
            let (Some(before), Some(after)) = (&before.web, &after.web) else {
                return Ok(false);
            };
            patch_bool(
                tree,
                node,
                "xml",
                before.xml_source,
                after.xml_source,
                edits,
            )?;
            patch_bool(
                tree,
                node,
                "sourceData",
                before.source_data,
                after.source_data,
                edits,
            )?;
            patch_bool(
                tree,
                node,
                "parsePre",
                before.parse_pre,
                after.parse_pre,
                edits,
            )?;
            patch_bool(
                tree,
                node,
                "consecutive",
                before.consecutive,
                after.consecutive,
                edits,
            )?;
            patch_bool(
                tree,
                node,
                "firstRow",
                before.first_row,
                after.first_row,
                edits,
            )?;
            patch_bool(tree, node, "xl97", before.excel97, after.excel97, edits)?;
            patch_bool(
                tree,
                node,
                "textDates",
                before.text_dates,
                after.text_dates,
                edits,
            )?;
            patch_bool(
                tree,
                node,
                "xl2000",
                before.excel2000,
                after.excel2000,
                edits,
            )?;
            patch_optional_attr(
                tree,
                node,
                "url",
                before.url.as_deref(),
                after.url.as_deref(),
                edits,
            )?;
            patch_optional_attr(
                tree,
                node,
                "post",
                before.post.as_deref(),
                after.post.as_deref(),
                edits,
            )?;
            patch_bool(
                tree,
                node,
                "htmlTables",
                before.html_tables,
                after.html_tables,
                edits,
            )?;
            patch_optional_attr(
                tree,
                node,
                "htmlFormat",
                before.html_format.map(html_str),
                after.html_format.map(html_str),
                edits,
            )?;
            patch_optional_attr(
                tree,
                node,
                "editPage",
                before.edit_page.as_deref(),
                after.edit_page.as_deref(),
                edits,
            )?;
            match (&before.tables, &after.tables) {
                (Some(before), Some(after)) => {
                    let collection = tree
                        .child(node, "tables")?
                        .ok_or_else(|| invalid("connection source is missing 'tables'"))?;
                    patch_sequence_source(
                        tree,
                        source,
                        collection,
                        before,
                        after,
                        canonical_web_table,
                        web_table_key,
                        patch_web_table_item,
                        edits,
                    )?;
                },
                (None, Some(after)) => {
                    let replacement = canonical_web_tables(after)?;
                    patch_optional_nested_collection(
                        tree,
                        source,
                        node,
                        "tables",
                        Some(&replacement),
                        edits,
                    )?;
                },
                (Some(_), None) => {
                    patch_optional_nested_collection(tree, source, node, "tables", None, edits)?;
                },
                (None, None) => {},
            }
            Ok(true)
        },
        "textPr" => {
            let (Some(before), Some(after)) = (&before.text, &after.text) else {
                return Ok(false);
            };
            patch_bool(tree, node, "prompt", before.prompt, after.prompt, edits)?;
            patch_optional_attr(
                tree,
                node,
                "fileType",
                before.file_type.map(file_str),
                after.file_type.map(file_str),
                edits,
            )?;
            patch_number(
                tree,
                node,
                "codePage",
                before.code_page,
                after.code_page,
                edits,
            )?;
            patch_optional_attr(
                tree,
                node,
                "characterSet",
                before.character_set.as_deref(),
                after.character_set.as_deref(),
                edits,
            )?;
            patch_number(
                tree,
                node,
                "firstRow",
                before.first_row,
                after.first_row,
                edits,
            )?;
            patch_optional_attr(
                tree,
                node,
                "sourceFile",
                before.source_file.as_deref(),
                after.source_file.as_deref(),
                edits,
            )?;
            patch_bool(
                tree,
                node,
                "delimited",
                before.delimited,
                after.delimited,
                edits,
            )?;
            patch_optional_attr(
                tree,
                node,
                "decimal",
                before.decimal.as_deref(),
                after.decimal.as_deref(),
                edits,
            )?;
            patch_optional_attr(
                tree,
                node,
                "thousands",
                before.thousands.as_deref(),
                after.thousands.as_deref(),
                edits,
            )?;
            patch_bool(tree, node, "tab", before.tab, after.tab, edits)?;
            patch_bool(tree, node, "space", before.space, after.space, edits)?;
            patch_bool(tree, node, "comma", before.comma, after.comma, edits)?;
            patch_bool(
                tree,
                node,
                "semicolon",
                before.semicolon,
                after.semicolon,
                edits,
            )?;
            patch_bool(
                tree,
                node,
                "consecutive",
                before.consecutive,
                after.consecutive,
                edits,
            )?;
            patch_optional_attr(
                tree,
                node,
                "qualifier",
                before.qualifier.map(qualifier_str),
                after.qualifier.map(qualifier_str),
                edits,
            )?;
            patch_optional_attr(
                tree,
                node,
                "delimiter",
                before.delimiter.as_deref(),
                after.delimiter.as_deref(),
                edits,
            )?;
            match (&before.fields, &after.fields) {
                (Some(before), Some(after)) => {
                    let collection = tree
                        .child(node, "textFields")?
                        .ok_or_else(|| invalid("connection source is missing 'textFields'"))?;
                    patch_sequence_source(
                        tree,
                        source,
                        collection,
                        before,
                        after,
                        canonical_text_field,
                        text_field_key,
                        patch_text_field_item,
                        edits,
                    )?;
                },
                (None, Some(after)) => {
                    let replacement = canonical_text_fields(after)?;
                    patch_optional_nested_collection(
                        tree,
                        source,
                        node,
                        "textFields",
                        Some(&replacement),
                        edits,
                    )?;
                },
                (Some(_), None) => {
                    patch_optional_nested_collection(
                        tree,
                        source,
                        node,
                        "textFields",
                        None,
                        edits,
                    )?;
                },
                (None, None) => {},
            }
            Ok(true)
        },
        "parameters" => {
            let (Some(before), Some(after)) = (&before.parameters, &after.parameters) else {
                return Ok(false);
            };
            patch_sequence_source(
                tree,
                source,
                node,
                before,
                after,
                canonical_parameter,
                parameter_key,
                patch_parameter_item,
                edits,
            )?;
            Ok(true)
        },
        "extLst" => Ok(false),
        _ => Ok(false),
    }
}

fn canonical_web_table(value: &WebTableSelector) -> Result<Vec<u8>> {
    let mut output = BoundedXml::new();
    match value {
        WebTableSelector::Missing => output.push_str("<m/>")?,
        WebTableSelector::String(value) => {
            output.push_str("<s")?;
            attr(&mut output, "v", value)?;
            output.push_str("/>")?;
        },
        WebTableSelector::Index(value) => {
            output.push_str("<x")?;
            num(&mut output, "v", *value)?;
            output.push_str("/>")?;
        },
    }
    Ok(output.finish())
}

fn web_table_value(value: &WebTableSelector) -> Option<String> {
    match value {
        WebTableSelector::Missing => None,
        WebTableSelector::String(value) => Some(value.clone()),
        WebTableSelector::Index(value) => Some(value.to_string()),
    }
}

fn web_table_key(value: &WebTableSelector) -> u64 {
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    match value {
        WebTableSelector::Missing => hasher.write_u8(0),
        WebTableSelector::String(value) => {
            hasher.write_u8(1);
            value.hash(&mut hasher);
        },
        WebTableSelector::Index(value) => {
            hasher.write_u8(2);
            value.hash(&mut hasher);
        },
    }
    hasher.finish()
}

fn canonical_web_tables(values: &[WebTableSelector]) -> Result<Vec<u8>> {
    let mut output = BoundedXml::new();
    output.push_str("<tables")?;
    num(&mut output, "count", values.len())?;
    output.push_char('>')?;
    for value in values {
        output.push_bytes(&canonical_web_table(value)?)?;
    }
    output.push_str("</tables>")?;
    Ok(output.finish())
}

fn canonical_text_field(value: &TextField) -> Result<Vec<u8>> {
    let mut output = BoundedXml::new();
    output.push_str("<textField")?;
    if let Some(value) = value.field_type {
        attr(&mut output, "type", field_str(value))?;
    }
    num_opt(&mut output, "position", value.position)?;
    output.push_str("/>")?;
    Ok(output.finish())
}

fn text_field_key(value: &TextField) -> u64 {
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    value.field_type.map(field_str).hash(&mut hasher);
    value.position.hash(&mut hasher);
    hasher.finish()
}

fn canonical_text_fields(values: &[TextField]) -> Result<Vec<u8>> {
    let mut output = BoundedXml::new();
    output.push_str("<textFields")?;
    num(&mut output, "count", values.len())?;
    output.push_char('>')?;
    for value in values {
        output.push_bytes(&canonical_text_field(value)?)?;
    }
    output.push_str("</textFields>")?;
    Ok(output.finish())
}

fn canonical_parameter(value: &ConnectionParameter) -> Result<Vec<u8>> {
    let mut output = BoundedXml::new();
    output.push_str("<parameter")?;
    str_opt(&mut output, "name", value.name.as_deref())?;
    num_opt(&mut output, "sqlType", value.sql_type)?;
    if let Some(parameter_type) = value.parameter_type {
        attr(&mut output, "parameterType", parameter_str(parameter_type))?;
    }
    bool_opt(&mut output, "refreshOnChange", value.refresh_on_change)?;
    str_opt(&mut output, "prompt", value.prompt.as_deref())?;
    bool_opt(&mut output, "boolean", value.boolean)?;
    if let Some(double) = value.double {
        let value = double.to_string();
        attr(&mut output, "double", &value)?;
    }
    num_opt(&mut output, "integer", value.integer)?;
    str_opt(&mut output, "string", value.string.as_deref())?;
    str_opt(&mut output, "cell", value.cell.as_deref())?;
    output.push_str("/>")?;
    Ok(output.finish())
}

fn parameter_key(value: &ConnectionParameter) -> u64 {
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    value.name.hash(&mut hasher);
    value.sql_type.hash(&mut hasher);
    value.parameter_type.map(parameter_str).hash(&mut hasher);
    value.refresh_on_change.hash(&mut hasher);
    value.prompt.hash(&mut hasher);
    value.boolean.hash(&mut hasher);
    value.double.map(f64::to_bits).hash(&mut hasher);
    value.integer.hash(&mut hasher);
    value.string.hash(&mut hasher);
    value.cell.hash(&mut hasher);
    hasher.finish()
}

fn patch_source_local_name(
    tree: &SourceTree,
    source: &[u8],
    node: usize,
    local: &str,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    let start = tree.nodes[node]
        .start
        .checked_add(1)
        .ok_or_else(|| invalid("connection source element-name offset overflowed"))?;
    let end = source_name_end(source, start);
    let qualified = std::str::from_utf8(
        source
            .get(start..end)
            .ok_or_else(|| invalid("connection source element-name span is invalid"))?,
    )
    .map_err(xml_error)?;
    let (prefix, old_local) = source_name_parts(qualified)?;
    if old_local != tree.nodes[node].local {
        return Err(invalid(
            "connection source element name changed during patch",
        ));
    }
    let local_start = start + prefix.len() + usize::from(!prefix.is_empty());
    edits.push(SourceEdit {
        range: local_start..end,
        replacement: local.as_bytes().to_vec(),
    });
    if !tree.nodes[node].self_closing {
        let close_start = tree.nodes[node]
            .end_start
            .checked_add(2)
            .ok_or_else(|| invalid("connection source closing-name offset overflowed"))?;
        let close_end = source_name_end(source, close_start);
        let closing = std::str::from_utf8(
            source
                .get(close_start..close_end)
                .ok_or_else(|| invalid("connection source closing-name span is invalid"))?,
        )
        .map_err(xml_error)?;
        let (close_prefix, close_local) = source_name_parts(closing)?;
        if close_local != tree.nodes[node].local || close_prefix != prefix {
            return Err(invalid(
                "connection source closing name does not match element",
            ));
        }
        let close_local_start =
            close_start + close_prefix.len() + usize::from(!close_prefix.is_empty());
        edits.push(SourceEdit {
            range: close_local_start..close_end,
            replacement: local.as_bytes().to_vec(),
        });
    }
    Ok(())
}

fn patch_web_table_item(
    tree: &SourceTree,
    source: &[u8],
    node: usize,
    before: &WebTableSelector,
    after: &WebTableSelector,
    edits: &mut Vec<SourceEdit>,
) -> Result<bool> {
    let expected_before = match before {
        WebTableSelector::Missing => "m",
        WebTableSelector::String(_) => "s",
        WebTableSelector::Index(_) => "x",
    };
    if tree.nodes[node].local != expected_before {
        return Ok(false);
    }
    let expected_after = match after {
        WebTableSelector::Missing => "m",
        WebTableSelector::String(_) => "s",
        WebTableSelector::Index(_) => "x",
    };
    if expected_before != expected_after {
        patch_source_local_name(tree, source, node, expected_after, edits)?;
    }
    let before_value = web_table_value(before);
    let after_value = web_table_value(after);
    patch_optional_attr(
        tree,
        node,
        "v",
        before_value.as_deref(),
        after_value.as_deref(),
        edits,
    )?;
    Ok(true)
}

fn patch_text_field_item(
    tree: &SourceTree,
    _source: &[u8],
    node: usize,
    before: &TextField,
    after: &TextField,
    edits: &mut Vec<SourceEdit>,
) -> Result<bool> {
    if tree.nodes[node].local != "textField" {
        return Ok(false);
    }
    patch_optional_attr(
        tree,
        node,
        "type",
        before.field_type.map(field_str),
        after.field_type.map(field_str),
        edits,
    )?;
    patch_number(
        tree,
        node,
        "position",
        before.position,
        after.position,
        edits,
    )?;
    Ok(true)
}

fn patch_parameter_item(
    tree: &SourceTree,
    _source: &[u8],
    node: usize,
    before: &ConnectionParameter,
    after: &ConnectionParameter,
    edits: &mut Vec<SourceEdit>,
) -> Result<bool> {
    if tree.nodes[node].local != "parameter" {
        return Ok(false);
    }
    patch_optional_attr(
        tree,
        node,
        "name",
        before.name.as_deref(),
        after.name.as_deref(),
        edits,
    )?;
    patch_number(
        tree,
        node,
        "sqlType",
        before.sql_type,
        after.sql_type,
        edits,
    )?;
    patch_optional_attr(
        tree,
        node,
        "parameterType",
        before.parameter_type.map(parameter_str),
        after.parameter_type.map(parameter_str),
        edits,
    )?;
    patch_bool(
        tree,
        node,
        "refreshOnChange",
        before.refresh_on_change,
        after.refresh_on_change,
        edits,
    )?;
    patch_optional_attr(
        tree,
        node,
        "prompt",
        before.prompt.as_deref(),
        after.prompt.as_deref(),
        edits,
    )?;
    patch_bool(tree, node, "boolean", before.boolean, after.boolean, edits)?;
    let before_double = before.double.map(|value| value.to_string());
    let after_double = after.double.map(|value| value.to_string());
    patch_optional_attr(
        tree,
        node,
        "double",
        before_double.as_deref(),
        after_double.as_deref(),
        edits,
    )?;
    patch_number(tree, node, "integer", before.integer, after.integer, edits)?;
    patch_optional_attr(
        tree,
        node,
        "string",
        before.string.as_deref(),
        after.string.as_deref(),
        edits,
    )?;
    patch_optional_attr(
        tree,
        node,
        "cell",
        before.cell.as_deref(),
        after.cell.as_deref(),
        edits,
    )?;
    Ok(true)
}

fn patch_optional_nested_collection(
    tree: &SourceTree,
    source: &[u8],
    parent: usize,
    name: &str,
    after: Option<&[u8]>,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    let existing = tree.child(parent, name)?;
    match (existing, after) {
        (Some(node), None) => {
            // Removing this selected collection is an explicit semantic
            // deletion. Its own descendants, comments, PIs, and inactive
            // MCE branch are part of that selected subtree and may go with it;
            // unrelated siblings remain source-bound.
            edits.push(SourceEdit {
                range: tree.nodes[node].start..tree.nodes[node].end,
                replacement: Vec::new(),
            });
        },
        (None, Some(replacement)) => {
            let insertion_parent = tree
                .semantic_insertion_parent(parent, name)
                .unwrap_or(parent);
            if tree.nodes[insertion_parent].self_closing {
                expand_self_closing_nested(tree, source, insertion_parent, replacement, edits)?;
            } else {
                let offset = child_insertion_offset(
                    tree,
                    insertion_parent,
                    tree.nodes[parent].namespace.as_ref(),
                    name,
                );
                edits.push(SourceEdit {
                    range: offset..offset,
                    replacement: qualify_fragment(replacement.to_vec(), tree, insertion_parent)?,
                });
            }
        },
        (Some(_), Some(_)) | (None, None) => {},
    }
    Ok(())
}

fn patch_sequence_source<T, Make, Patch>(
    tree: &SourceTree,
    source: &[u8],
    collection: usize,
    before: &[T],
    after: &[T],
    make_item: Make,
    item_key: impl Fn(&T) -> u64,
    mut patch_item: Patch,
    edits: &mut Vec<SourceEdit>,
) -> Result<()>
where
    T: PartialEq,
    Make: Fn(&T) -> Result<Vec<u8>>,
    Patch: FnMut(&SourceTree, &[u8], usize, &T, &T, &mut Vec<SourceEdit>) -> Result<bool>,
{
    let source_nodes = tree.semantic_elements(collection)?;
    if source_nodes.len() != before.len() {
        return Err(invalid(
            "connection source has an ambiguous nested collection layout",
        ));
    }
    let source_parent = source_nodes
        .first()
        .and_then(|node| tree.nodes[*node].parent)
        .unwrap_or(collection);
    let multiple_source_parents = source_nodes
        .iter()
        .any(|node| tree.nodes[*node].parent != Some(source_parent));

    let mut exact = HashMap::<u64, Vec<usize>>::new();
    exact
        .try_reserve(before.len())
        .map_err(|_| invalid("connection nested collection matching allocation failed"))?;
    for (index, value) in before.iter().enumerate() {
        let bucket = exact.entry(item_key(value)).or_default();
        bucket
            .try_reserve(1)
            .map_err(|_| invalid("connection nested collection matching allocation failed"))?;
        bucket.push(index);
    }
    let mut source_for_target = Vec::new();
    source_for_target
        .try_reserve_exact(after.len())
        .map_err(|_| invalid("connection nested collection matching allocation failed"))?;
    source_for_target.resize(after.len(), None);
    let mut used = Vec::new();
    used.try_reserve_exact(before.len())
        .map_err(|_| invalid("connection nested collection matching allocation failed"))?;
    used.resize(before.len(), false);
    for (target, value) in after.iter().enumerate() {
        let key = item_key(value);
        let Some(bucket) = exact.get_mut(&key) else {
            continue;
        };
        if let Some(bucket_index) = bucket.iter().position(|source| before[*source] == *value) {
            let source = bucket.swap_remove(bucket_index);
            used[source] = true;
            source_for_target[target] = Some(source);
        }
    }
    let mut next_unused = 0;
    for target in 0..after.len() {
        if source_for_target[target].is_some() {
            continue;
        }
        let source = if target < before.len() && !used[target] {
            Some(target)
        } else {
            while next_unused < used.len() && used[next_unused] {
                next_unused += 1;
            }
            (next_unused < used.len()).then_some(next_unused)
        };
        if let Some(source) = source {
            used[source] = true;
            source_for_target[target] = Some(source);
        }
    }
    if multiple_source_parents {
        if before.len() != after.len() {
            return Err(invalid(
                "connection nested collection has ambiguous cross-branch insertion or deletion",
            ));
        }
        for (target, source_index) in source_for_target.iter().enumerate() {
            let source_index = (*source_index).ok_or_else(|| {
                invalid("connection nested collection has an ambiguous cross-branch edit")
            })?;
            if tree.nodes[source_nodes[source_index]].parent
                != tree.nodes[source_nodes[target]].parent
            {
                return Err(invalid(
                    "connection nested collection move crosses source structural branches",
                ));
            }
        }
    }
    for (target, value) in after.iter().enumerate() {
        let range = if target < source_nodes.len() {
            tree.nodes[source_nodes[target]].start..tree.nodes[source_nodes[target]].end
        } else {
            tree.nodes[source_parent].end_start..tree.nodes[source_parent].end_start
        };
        match source_for_target[target] {
            Some(source_index) if source_index == target => {
                if before[source_index] != *value
                    && !patch_item(
                        tree,
                        source,
                        source_nodes[source_index],
                        &before[source_index],
                        value,
                        edits,
                    )?
                {
                    return Err(invalid(
                        "connection nested item edit cannot preserve source markup",
                    ));
                }
            },
            Some(source_index) => {
                let source_node = source_nodes[source_index];
                let span = source
                    .get(tree.nodes[source_node].start..tree.nodes[source_node].end)
                    .ok_or_else(|| invalid("connection nested item source span is invalid"))?
                    .to_vec();
                let replacement = if before[source_index] == *value {
                    span
                } else {
                    let mut item_edits = Vec::new();
                    if !patch_item(
                        tree,
                        source,
                        source_node,
                        &before[source_index],
                        value,
                        &mut item_edits,
                    )? {
                        return Err(invalid(
                            "connection nested item edit cannot preserve source markup",
                        ));
                    }
                    let start = tree.nodes[source_node].start;
                    let mut relative = Vec::new();
                    relative
                        .try_reserve_exact(item_edits.len())
                        .map_err(|_| invalid("connection nested item edit allocation failed"))?;
                    for edit in item_edits {
                        if edit.range.start < start || edit.range.end < start {
                            return Err(invalid(
                                "connection nested item edit escapes its source span",
                            ));
                        }
                        relative.push(SourceEdit {
                            range: (edit.range.start - start)..(edit.range.end - start),
                            replacement: edit.replacement,
                        });
                    }
                    apply_source_edits(&span, relative)?
                };
                edits.push(SourceEdit { range, replacement });
            },
            None => edits.push(SourceEdit {
                range,
                replacement: qualify_fragment(make_item(value)?, tree, source_parent)?,
            }),
        }
    }
    for &source_node in source_nodes.iter().take(before.len()).skip(after.len()) {
        edits.push(SourceEdit {
            range: tree.nodes[source_node].start..tree.nodes[source_node].end,
            replacement: Vec::new(),
        });
    }

    if before.len() != after.len() {
        let count = after.len().to_string();
        let source_count = tree.attribute(source, collection, "count")?;
        patch_optional_attr(
            tree,
            collection,
            "count",
            source_count.as_deref(),
            Some(&count),
            edits,
        )?;
    }
    Ok(())
}

fn expand_self_closing_nested(
    tree: &SourceTree,
    source: &[u8],
    parent: usize,
    child: &[u8],
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    if !tree.nodes[parent].self_closing {
        return Err(invalid(
            "connection nested child expansion has a non-self-closing parent",
        ));
    }
    let opening_start = tree.nodes[parent].start;
    let opening_end = tree.nodes[parent].close_pos;
    let mut opening_edits = Vec::new();
    let mut retained_edits = Vec::with_capacity(edits.len());
    for edit in edits.drain(..) {
        if edit.range.start >= opening_start && edit.range.end <= opening_end {
            opening_edits.push(SourceEdit {
                range: (edit.range.start - opening_start)..(edit.range.end - opening_start),
                replacement: edit.replacement,
            });
        } else {
            retained_edits.push(edit);
        }
    }
    *edits = retained_edits;
    let mut expanded = apply_source_edits(
        source
            .get(opening_start..opening_end)
            .ok_or_else(|| invalid("connection nested opening span is invalid"))?,
        opening_edits,
    )?;
    expanded.push(b'>');
    expanded.extend_from_slice(&qualify_fragment(child.to_vec(), tree, parent)?);
    expanded.extend_from_slice(b"</");
    expanded.extend_from_slice(tree.nodes[parent].qualified.as_bytes());
    expanded.push(b'>');
    edits.push(SourceEdit {
        range: tree.nodes[parent].start..tree.nodes[parent].end,
        replacement: expanded,
    });
    Ok(())
}

fn expand_self_closing(
    tree: &SourceTree,
    source: &[u8],
    parent: usize,
    mut children: Vec<(String, Vec<u8>)>,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    if children.is_empty() {
        return Ok(());
    }
    if !tree.nodes[parent].self_closing {
        return Err(invalid(
            "connection child expansion has a non-self-closing parent",
        ));
    }
    let opening_start = tree.nodes[parent].start;
    let opening_end = tree.nodes[parent].close_pos;
    let mut opening_edits = Vec::new();
    let mut retained_edits = Vec::with_capacity(edits.len());
    for edit in edits.drain(..) {
        if edit.range.start >= opening_start && edit.range.end <= opening_end {
            opening_edits.push(SourceEdit {
                range: (edit.range.start - opening_start)..(edit.range.end - opening_start),
                replacement: edit.replacement,
            });
        } else {
            retained_edits.push(edit);
        }
    }
    *edits = retained_edits;
    children.sort_by_key(|(name, _replacement)| child_order(name));
    let mut expanded = apply_source_edits(
        source
            .get(opening_start..opening_end)
            .ok_or_else(|| invalid("connection source opening span is invalid"))?,
        opening_edits,
    )?;
    expanded.push(b'>');
    for (_name, replacement) in children {
        expanded.extend_from_slice(&qualify_fragment(replacement, tree, parent)?);
    }
    expanded.extend_from_slice(b"</");
    expanded.extend_from_slice(tree.nodes[parent].qualified.as_bytes());
    expanded.push(b'>');
    edits.push(SourceEdit {
        range: tree.nodes[parent].start..tree.nodes[parent].end,
        replacement: expanded,
    });
    Ok(())
}

fn child_order(name: &str) -> u8 {
    match name {
        "dbPr" => 0,
        "olapPr" => 1,
        "webPr" => 2,
        "textPr" => 3,
        "parameters" => 4,
        "extLst" => 5,
        _ => 6,
    }
}

fn child_insertion_offset(tree: &SourceTree, parent: usize, namespace: &str, name: &str) -> usize {
    let order = child_order(name);
    tree.nodes[parent]
        .children
        .iter()
        .copied()
        .find(|child| {
            tree.nodes[*child].namespace.as_ref() == namespace
                && child_order(&tree.nodes[*child].local) > order
        })
        .map(|child| tree.nodes[child].start)
        .unwrap_or(tree.nodes[parent].end_start)
}

fn qualify_fragment(mut fragment: Vec<u8>, tree: &SourceTree, parent: usize) -> Result<Vec<u8>> {
    let namespace = tree.nodes[tree.root].namespace.as_ref();
    if tree
        .nodes
        .get(parent)
        .and_then(|node| node.context.resolve_uri(""))
        .is_some_and(|binding| binding.as_ref() == namespace)
    {
        return Ok(fragment);
    }
    let end = source_tag_end(&fragment, 0)?;
    if fragment_has_default_namespace_declaration(&fragment, end)? {
        return Ok(fragment);
    }
    let insert = fragment[..end]
        .iter()
        .rposition(|byte| !byte.is_ascii_whitespace())
        .filter(|position| fragment[*position] == b'/')
        .unwrap_or(end);
    let mut declaration = b" xmlns=\"".to_vec();
    declaration.extend_from_slice(namespace.as_bytes());
    declaration.push(b'"');
    fragment.splice(insert..insert, declaration);
    Ok(fragment)
}

fn fragment_has_default_namespace_declaration(fragment: &[u8], end: usize) -> Result<bool> {
    let mut reader = Reader::from_reader(
        fragment
            .get(..=end)
            .ok_or_else(|| invalid("connection fragment opening span is invalid"))?,
    );
    match reader.read_event() {
        Ok(Event::Start(element) | Event::Empty(element)) => {
            for attribute in element.attributes().with_checks(true) {
                if attribute.map_err(xml_error)?.key.as_ref() == b"xmlns" {
                    return Ok(true);
                }
            }
            Ok(false)
        },
        Ok(_) => Err(invalid(
            "connection fragment does not start with an element",
        )),
        Err(error) => Err(xml_error(error)),
    }
}

fn patch_bool(
    tree: &SourceTree,
    node: usize,
    name: &str,
    before: Option<bool>,
    after: Option<bool>,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    patch_optional_attr(
        tree,
        node,
        name,
        before.map(|value| if value { "1" } else { "0" }),
        after.map(|value| if value { "1" } else { "0" }),
        edits,
    )
}

fn patch_number<T: std::fmt::Display + Copy>(
    tree: &SourceTree,
    node: usize,
    name: &str,
    before: Option<T>,
    after: Option<T>,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    let before = before.map(|value| value.to_string());
    let after = after.map(|value| value.to_string());
    patch_optional_attr(tree, node, name, before.as_deref(), after.as_deref(), edits)
}

fn patch_optional_attr(
    tree: &SourceTree,
    node: usize,
    name: &str,
    before: Option<&str>,
    after: Option<&str>,
    edits: &mut Vec<SourceEdit>,
) -> Result<()> {
    if before == after {
        return Ok(());
    }
    let attribute = tree.nodes[node].attrs.iter().find(|attribute| {
        !attribute.namespace_declaration
            && attribute.namespace.is_empty()
            && attribute.local == name
    });
    match (attribute, after) {
        (Some(attribute), Some(value)) => edits.push(SourceEdit {
            range: attribute.value_start..attribute.value_end,
            replacement: escape_attribute(value),
        }),
        (Some(attribute), None) => edits.push(SourceEdit {
            range: attribute.start..attribute.value_end + 1,
            replacement: Vec::new(),
        }),
        (None, Some(value)) if before.is_none() => edits.push(SourceEdit {
            range: tree.nodes[node].close_pos..tree.nodes[node].close_pos,
            replacement: format!(
                " {name}=\"{}\"",
                String::from_utf8_lossy(&escape_attribute(value))
            )
            .into_bytes(),
        }),
        (None, Some(_)) => {
            return Err(invalid(format!(
                "connection source is missing attribute '{name}'"
            )));
        },
        (None, None) => {
            if before.is_some() {
                return Err(invalid(format!(
                    "connection source is missing attribute '{name}'"
                )));
            }
        },
    }
    Ok(())
}

fn escape_attribute(value: &str) -> Vec<u8> {
    let mut result = Vec::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => result.extend_from_slice(b"&amp;"),
            '<' => result.extend_from_slice(b"&lt;"),
            '"' => result.extend_from_slice(b"&quot;"),
            '\r' => result.extend_from_slice(b"&#xD;"),
            '\n' => result.extend_from_slice(b"&#xA;"),
            '\t' => result.extend_from_slice(b"&#x9;"),
            _ => {
                let mut encoded = [0; 4];
                result.extend_from_slice(character.encode_utf8(&mut encoded).as_bytes());
            },
        }
    }
    result
}

#[derive(Debug)]
struct SourceEdit {
    range: Range<usize>,
    replacement: Vec<u8>,
}

fn apply_source_edits(source: &[u8], mut edits: Vec<SourceEdit>) -> Result<Vec<u8>> {
    let mut indexed = edits.drain(..).enumerate().collect::<Vec<_>>();
    indexed.sort_by(|(left_index, left), (right_index, right)| {
        left.range
            .start
            .cmp(&right.range.start)
            .then_with(|| left_index.cmp(right_index))
    });
    let mut output_len = source.len();
    let mut covered_end = 0;
    for (_index, edit) in &indexed {
        if edit.range.start > edit.range.end || edit.range.end > source.len() {
            return Err(invalid("connection source edit is out of bounds"));
        }
        if edit.range.start < covered_end {
            return Err(invalid("connection source edits overlap"));
        }
        output_len = output_len
            .checked_sub(edit.range.end - edit.range.start)
            .and_then(|length| length.checked_add(edit.replacement.len()))
            .ok_or_else(|| invalid("connection source output length overflows"))?;
        covered_end = edit.range.end;
    }
    if output_len > MAX_XML_BYTES {
        return Err(invalid("connection source output exceeds 16 MiB"));
    }
    let mut result = Vec::new();
    result
        .try_reserve_exact(output_len)
        .map_err(|_source| invalid("connection source output allocation failed"))?;
    let mut cursor = 0;
    for (_index, edit) in indexed {
        result.extend_from_slice(&source[cursor..edit.range.start]);
        result.extend_from_slice(&edit.replacement);
        cursor = edit.range.end;
    }
    result.extend_from_slice(&source[cursor..]);
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn connection(id: u32, name: Option<&str>) -> Connection {
        Connection {
            id,
            source_file: None,
            odc_file: None,
            keep_alive: None,
            interval: None,
            name: name.map(str::to_owned),
            description: None,
            connection_type: None,
            reconnection_method: None,
            refreshed_version: 7,
            min_refreshable_version: None,
            save_password: None,
            new_connection: None,
            deleted: None,
            only_use_connection_file: None,
            background: None,
            refresh_on_load: None,
            save_data: None,
            credentials: None,
            single_sign_on_id: None,
            database: None,
            olap: None,
            web: None,
            text: None,
            parameters: None,
            extension_xml: None,
        }
    }

    fn database_connection() -> DatabaseProperties {
        DatabaseProperties {
            connection: "Provider=opaque".into(),
            command: None,
            server_command: None,
            command_type: None,
        }
    }

    #[test]
    fn processing_instruction_filter_borrows_clean_input_and_splices_exact_spans() {
        let clean = b"<root><child/></root>";
        let filtered = strip_processing_instructions(clean).expect("clean XML");
        assert!(matches!(filtered, Cow::Borrowed(_)));

        let source = b"<root a=\"1\"><?one?><child/>\n<?two data?>tail</root>";
        let filtered = strip_processing_instructions(source).expect("XML with PIs");
        assert!(matches!(&filtered, Cow::Owned(_)));
        assert_eq!(filtered.as_ref(), b"<root a=\"1\"><child/>\ntail</root>");
    }

    #[test]
    fn source_tree_shares_inherited_namespace_contexts_and_uri_storage() {
        let large_uri = format!("urn:large:{}", "x".repeat(64 * 1024));
        let mut source = format!(
            r#"<connections xmlns="{CORE_NAMESPACE}" xmlns:f="{large_uri}"><connection id="1" refreshedVersion="7" name="before"/>"#
        );
        for _ in 0..512 {
            source.push_str("<f:future/>");
        }
        source.push_str(r#"<f:holder xmlns:f="urn:shadow"><f:item/></f:holder></connections>"#);

        let tree = SourceTree::parse(source.as_bytes()).unwrap();
        let root_uri = tree.nodes[tree.root]
            .context
            .resolve_uri("f")
            .cloned()
            .unwrap();
        let large_nodes = tree
            .nodes
            .iter()
            .filter(|node| node.local == "future")
            .collect::<Vec<_>>();
        assert_eq!(large_nodes.len(), 512);
        for node in &large_nodes {
            assert!(Arc::ptr_eq(
                node.context.resolve_uri("f").unwrap(),
                &root_uri
            ));
            assert!(Arc::ptr_eq(&node.namespace, &root_uri));
        }
        let shadowed = tree.nodes.iter().find(|node| node.local == "item").unwrap();
        assert_eq!(shadowed.namespace.as_ref(), "urn:shadow");
        assert!(!Arc::ptr_eq(&shadowed.namespace, &large_nodes[0].namespace));

        let before = Connections::parse(source.as_bytes()).unwrap();
        let mut after = before.clone();
        after.connections[0].name = Some("after".into());
        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let expected = source.replace("name=\"before\"", "name=\"after\"");
        assert_eq!(patched, expected.as_bytes());
        assert_eq!(Connections::parse(&patched).unwrap(), after);
    }

    #[test]
    fn parse_dom_shares_inherited_namespace_contexts_and_uri_storage() {
        let large_uri = format!("urn:large:{}", "x".repeat(64 * 1024));
        let mut source = format!(
            r#"<connections xmlns="{CORE_NAMESPACE}" xmlns:f="{large_uri}"><connection id="1" refreshedVersion="7"/>"#
        );
        for _ in 0..512 {
            source.push_str("<f:future/>");
        }
        source.push_str(r#"<f:holder xmlns:f="urn:shadow"><f:item/></f:holder></connections>"#);

        let root = parse_dom(source.as_bytes()).unwrap();
        let root_uri = root.context.resolve_uri("f").cloned().unwrap();
        let large_nodes = root
            .content
            .iter()
            .filter_map(|content| match content {
                Content::Node(node) if node.l == "future" => Some(node),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(large_nodes.len(), 512);
        for node in &large_nodes {
            assert!(Arc::ptr_eq(
                node.context.resolve_uri("f").unwrap(),
                &root_uri
            ));
            assert!(Arc::ptr_eq(&node.ns, &root_uri));
        }
        let shadowed = root
            .content
            .iter()
            .find_map(|content| match content {
                Content::Node(node) if node.l == "holder" => {
                    node.content.iter().find_map(|child| match child {
                        Content::Node(child) if child.l == "item" => Some(child),
                        _ => None,
                    })
                },
                _ => None,
            })
            .unwrap();
        assert_eq!(shadowed.ns.as_ref(), "urn:shadow");
        assert!(!Arc::ptr_eq(&shadowed.ns, &large_nodes[0].ns));
    }

    #[test]
    fn namespace_context_rejects_invalid_reserved_bindings() {
        for source in [
            format!(
                r#"<connections xmlns="{CORE_NAMESPACE}" xmlns:xml="urn:invalid"><connection id="1" refreshedVersion="7"/></connections>"#
            ),
            format!(
                r#"<connections xmlns="{CORE_NAMESPACE}" xmlns:bad="http://www.w3.org/XML/1998/namespace"><connection id="1" refreshedVersion="7"/></connections>"#
            ),
            format!(
                r#"<connections xmlns="{CORE_NAMESPACE}" xmlns:bad="http://www.w3.org/2000/xmlns/"><connection id="1" refreshedVersion="7"/></connections>"#
            ),
        ] {
            assert!(SourceTree::parse(source.as_bytes()).is_err());
            assert!(parse_dom(source.as_bytes()).is_err());
        }
    }

    #[test]
    fn source_patch_resolves_prefixed_core_and_ignores_foreign_local_names() {
        let source = format!(
            r#"<p:connections xmlns:p="{CORE_NAMESPACE}" xmlns:f="urn:foreign"><p:connection f:id="99" id="&#49;" refreshedVersion="7" name="before"><f:dbPr f:marker="keep"/></p:connection></p:connections>"#
        );
        let before = Connections::parse(source.as_bytes()).unwrap();
        assert_eq!(before.connections[0].id, 1);
        let mut after = before.clone();
        after.connections[0].name = Some("after".into());
        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let text = std::str::from_utf8(&patched).unwrap();
        assert!(text.contains("p:connections"));
        assert!(text.contains("f:id=\"99\""));
        assert!(text.contains("f:dbPr f:marker=\"keep\""));
        assert!(!text.contains("<connection id="));
        assert_eq!(Connections::parse(&patched).unwrap(), after);
    }

    #[test]
    fn source_patch_inserts_children_in_schema_order_and_expands_self_closing_nodes() {
        let source = format!(
            r#"<p:connections xmlns:p="{CORE_NAMESPACE}"><p:connection id="1" refreshedVersion="7"><p:extLst/></p:connection></p:connections>"#
        );
        let before = Connections::parse(source.as_bytes()).unwrap();
        let mut after = before.clone();
        after.connections[0].database = Some(database_connection());
        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let text = std::str::from_utf8(&patched).unwrap();
        assert!(text.find("dbPr").unwrap() < text.find("extLst").unwrap());
        assert_eq!(Connections::parse(&patched).unwrap(), after);

        let source = format!(
            r#"<p:connections xmlns:p="{CORE_NAMESPACE}"><p:connection id="1" refreshedVersion="7"/></p:connections>"#
        );
        let before = Connections::parse(source.as_bytes()).unwrap();
        let mut after = before.clone();
        after.connections[0].database = Some(database_connection());
        after.connections[0].olap = Some(OlapProperties::default());
        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let text = std::str::from_utf8(&patched).unwrap();
        assert!(text.contains("</p:connection>"));
        assert!(text.find("dbPr").unwrap() < text.find("olapPr").unwrap());
        assert_eq!(Connections::parse(&patched).unwrap(), after);
    }

    #[test]
    fn source_patch_combines_scalar_and_child_changes_for_prefixed_and_unprefixed_self_closing_nodes()
     {
        let cases = [
            (
                format!(
                    r#"<connections xmlns="{CORE_NAMESPACE}" xmlns:f="urn:foreign"><connection id="1" refreshedVersion="7" name="before" f:marker="keep"/></connections>"#
                ),
                "</connection>",
            ),
            (
                format!(
                    r#"<p:connections xmlns:p="{CORE_NAMESPACE}" xmlns:f="urn:foreign"><p:connection id="1" refreshedVersion="7" name="before" f:marker="keep"/></p:connections>"#
                ),
                "</p:connection>",
            ),
        ];

        for (source, closing_tag) in cases {
            let before = Connections::parse(source.as_bytes()).unwrap();
            let mut after = before.clone();
            after.connections[0].name = Some("after".into());
            after.connections[0].database = Some(database_connection());

            let patched =
                patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
            let text = std::str::from_utf8(&patched).unwrap();
            assert!(text.contains(closing_tag));
            assert!(text.contains("name=\"after\""));
            assert!(text.contains("f:marker=\"keep\""));
            assert_eq!(Connections::parse(&patched).unwrap(), after);
        }
    }

    #[test]
    fn source_patch_preserves_comments_and_processing_instructions_in_replaced_children() {
        let source = format!(
            r#"<connections xmlns="{CORE_NAMESPACE}"><connection id="1" refreshedVersion="7"><dbPr connection="before"><?keep?><!--keep--></dbPr></connection></connections>"#
        );
        let before = Connections::parse(source.as_bytes()).unwrap();
        let mut after = before.clone();
        after.connections[0].database.as_mut().unwrap().connection = "after".into();

        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let text = std::str::from_utf8(&patched).unwrap();
        assert!(text.contains("<?keep?>"));
        assert!(text.contains("<!--keep-->"));
        assert!(text.contains("connection=\"after\""));
        assert_eq!(Connections::parse(&patched).unwrap(), after);
    }

    #[test]
    fn source_patch_replaces_extlst_as_an_explicit_whole_owner() {
        let source = format!(
            r#"<connections xmlns="{CORE_NAMESPACE}" xmlns:f="urn:foreign" xmlns:mc="{MCE}" xmlns:u="urn:unsupported"><connection id="1" refreshedVersion="7"><?outside-before?><extLst f:old="keep"><?old?><mc:AlternateContent><mc:Choice Requires="u"><u:old/></mc:Choice><mc:Fallback><!--old-fallback--><f:old-child/></mc:Fallback></mc:AlternateContent><!--old-comment--></extLst><?outside-after?></connection></connections>"#,
            MCE = litchi_ooxml_common::mce::NAMESPACE,
        );
        let before = Connections::parse(source.as_bytes()).unwrap();
        let mut after = before.clone();
        after.connections[0].extension_xml = Some(
            format!(
                r#"<extLst xmlns="{CORE_NAMESPACE}" xmlns:f="urn:foreign" xmlns:mc="{MCE}" xmlns:u="urn:unsupported" f:new="owned"><mc:AlternateContent><mc:Choice Requires="u"><u:new-inactive/></mc:Choice><mc:Fallback><f:new-child/></mc:Fallback></mc:AlternateContent><!--new-comment--></extLst>"#,
                CORE_NAMESPACE = CORE_NAMESPACE,
                MCE = litchi_ooxml_common::mce::NAMESPACE,
            )
            .into_bytes(),
        );

        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let text = std::str::from_utf8(&patched).unwrap();
        assert!(text.contains("<?outside-before?>") && text.contains("<?outside-after?>"));
        assert!(text.contains("f:new=\"owned\"") && text.contains("f:new-child"));
        assert!(text.contains("<!--new-comment-->") && text.contains("mc:AlternateContent"));
        assert!(!text.contains("f:old=\"keep\"") && !text.contains("<?old?>"));
        assert!(!text.contains("old-fallback") && !text.contains("old-comment"));
        Connections::parse(&patched).unwrap();

        let mut removed = before.clone();
        removed.connections[0].extension_xml = None;
        let patched =
            patch_connections_source(source.as_bytes(), &before, &removed, false).unwrap();
        let text = std::str::from_utf8(&patched).unwrap();
        assert!(text.contains("<?outside-before?>") && text.contains("<?outside-after?>"));
        assert!(!text.contains("<extLst") && !text.contains("<?old?>"));
        Connections::parse(&patched).unwrap();
    }

    #[test]
    fn source_patch_updates_web_scalar_without_rebuilding_mce_tables_or_pi_slots() {
        let source = format!(
            r#"<connections xmlns="{CORE_NAMESPACE}" xmlns:mc="{MCE}" xmlns:u="urn:unsupported"><connection id="1" refreshedVersion="7"><webPr url="before"><?before?><mc:AlternateContent><mc:Choice Requires="u"><u:tables count="1"><u:m/></u:tables></mc:Choice><mc:Fallback><?inside?><tables count="1"><m/></tables></mc:Fallback></mc:AlternateContent><?after?></webPr></connection></connections>"#,
            MCE = litchi_ooxml_common::mce::NAMESPACE,
        );
        let before = Connections::parse(source.as_bytes()).unwrap();
        let mut after = before.clone();
        after.connections[0].web.as_mut().unwrap().url = Some("after".into());

        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let expected = source.replace("url=\"before\"", "url=\"after\"");
        assert_eq!(patched, expected.as_bytes());
        assert_eq!(Connections::parse(&patched).unwrap(), after);
    }

    #[test]
    fn source_patch_updates_nested_collections_without_losing_source_slots() {
        let source = format!(
            r#"<connections xmlns="{CORE_NAMESPACE}"><connection id="1" refreshedVersion="7"><webPr><tables count="2"><?before?><m/><?inside?><s v="A"/></tables><?after?></webPr><textPr><textFields count="1"><?field?><textField type="text" position="1"/></textFields></textPr><parameters count="1"><parameter name="old"><?parameter?></parameter></parameters></connection></connections>"#
        );
        let before = Connections::parse(source.as_bytes()).unwrap();
        let mut after = before.clone();
        after.connections[0].web.as_mut().unwrap().tables = Some(vec![
            WebTableSelector::Missing,
            WebTableSelector::String("B".into()),
            WebTableSelector::Index(3),
        ]);
        after.connections[0].text.as_mut().unwrap().fields = Some(vec![
            TextField {
                field_type: Some(TextFieldType::Text),
                position: Some(1),
            },
            TextField {
                field_type: Some(TextFieldType::General),
                position: Some(2),
            },
        ]);
        after.connections[0]
            .parameters
            .as_mut()
            .unwrap()
            .push(ConnectionParameter {
                name: Some("new".into()),
                sql_type: None,
                parameter_type: None,
                refresh_on_change: None,
                prompt: None,
                boolean: None,
                double: None,
                integer: None,
                string: None,
                cell: None,
            });

        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let text = std::str::from_utf8(&patched).unwrap();
        for marker in [
            "<?before?>",
            "<?inside?>",
            "<?after?>",
            "<?field?>",
            "<?parameter?>",
        ] {
            assert!(text.contains(marker), "missing source marker {marker}");
        }
        assert!(text.contains("<tables count=\"3\">") && text.contains("<s v=\"B\"/>"));
        assert!(text.contains("<textFields count=\"2\">") && text.contains("position=\"2\""));
        assert!(text.contains("<parameters count=\"2\">") && text.contains("name=\"new\""));
        assert_eq!(Connections::parse(&patched).unwrap(), after);
    }

    #[test]
    fn source_patch_keeps_mce_branch_while_editing_nested_table_sequence() {
        let source = format!(
            r#"<connections xmlns="{CORE_NAMESPACE}" xmlns:mc="{MCE}" xmlns:u="urn:unsupported"><connection id="1" refreshedVersion="7"><webPr><mc:AlternateContent><mc:Choice Requires="u"><u:tables count="1"><u:s v="inactive"/></u:tables></mc:Choice><mc:Fallback><?inside?><tables count="1"><m/></tables></mc:Fallback></mc:AlternateContent></webPr></connection></connections>"#,
            MCE = litchi_ooxml_common::mce::NAMESPACE,
        );
        let before = Connections::parse(source.as_bytes()).unwrap();
        let mut after = before.clone();
        after.connections[0].web.as_mut().unwrap().tables = Some(vec![
            WebTableSelector::Missing,
            WebTableSelector::String("active".into()),
        ]);

        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let text = std::str::from_utf8(&patched).unwrap();
        assert!(text.contains("<mc:Choice") && text.contains("v=\"inactive\""));
        assert!(text.contains("<mc:Fallback>") && text.contains("<?inside?>"));
        assert!(text.contains("<tables count=\"2\">") && text.contains("<s v=\"active\"/>"));
        assert_eq!(Connections::parse(&patched).unwrap(), after);
    }

    #[test]
    fn source_patch_updates_same_index_items_across_direct_and_mce_parents() {
        let source = format!(
            r#"<connections xmlns="{CORE_NAMESPACE}" xmlns:mc="{MCE}" xmlns:u="urn:unsupported"><connection id="1" refreshedVersion="7"><webPr><tables count="2"><s v="A"/><mc:AlternateContent><mc:Choice Requires="u"><u:s v="inactive"/></mc:Choice><mc:Fallback><s v="B"/></mc:Fallback></mc:AlternateContent></tables></webPr></connection></connections>"#,
            MCE = litchi_ooxml_common::mce::NAMESPACE,
        );
        let before = Connections::parse(source.as_bytes()).unwrap();
        assert_eq!(
            before.connections[0].web.as_ref().unwrap().tables.as_ref(),
            Some(&vec![
                WebTableSelector::String("A".into()),
                WebTableSelector::String("B".into()),
            ])
        );
        let mut after = before.clone();
        after.connections[0].web.as_mut().unwrap().tables = Some(vec![
            WebTableSelector::String("C".into()),
            WebTableSelector::String("B".into()),
        ]);

        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let text = std::str::from_utf8(&patched).unwrap();
        assert!(text.contains("<s v=\"C\"/>") && text.contains("<s v=\"B\"/>"));
        assert!(text.contains(
            "<mc:AlternateContent><mc:Choice Requires=\"u\"><u:s v=\"inactive\"/></mc:Choice><mc:Fallback><s v=\"B\"/></mc:Fallback></mc:AlternateContent>"
        ));
        assert_eq!(Connections::parse(&patched).unwrap(), after);

        let mut moved = before.clone();
        moved.connections[0].web.as_mut().unwrap().tables = Some(vec![
            WebTableSelector::String("B".into()),
            WebTableSelector::String("A".into()),
        ]);
        let error =
            patch_connections_source(source.as_bytes(), &before, &moved, false).unwrap_err();
        assert!(
            error
                .to_string()
                .contains("move crosses source structural branches")
        );
    }

    #[test]
    fn source_patch_expands_self_closing_nested_collections_with_scalar_edits() {
        let source = format!(
            r#"<connections xmlns="{CORE_NAMESPACE}"><connection id="1" refreshedVersion="7"><webPr url="before"/></connection></connections>"#
        );
        let before = Connections::parse(source.as_bytes()).unwrap();
        let mut after = before.clone();
        let web = after.connections[0].web.as_mut().unwrap();
        web.url = Some("after".into());
        web.tables = Some(vec![WebTableSelector::String("A".into())]);

        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let text = std::str::from_utf8(&patched).unwrap();
        assert!(
            text.contains("<webPr url=\"after\"><tables count=\"1\"><s v=\"A\"/></tables></webPr>")
        );
        assert_eq!(Connections::parse(&patched).unwrap(), after);
    }

    #[test]
    fn source_patch_reorders_nested_items_without_losing_opaque_markup() {
        let source = format!(
            r#"<connections xmlns="{CORE_NAMESPACE}" xmlns:f="urn:f" xmlns:mc="{MCE}" xmlns:u="urn:u"><connection id="1" refreshedVersion="7"><webPr><tables count="2"><s v="A" xmlns:f="urn:f" f:opaque="table-a"><mc:AlternateContent><mc:Choice Requires="u"><u:inactive/></mc:Choice><mc:Fallback/></mc:AlternateContent></s><s v="B" xmlns:f="urn:f" f:opaque="table-b"/></tables></webPr><textPr><textFields count="2"><textField type="text" position="1" xmlns:f="urn:f" f:opaque="field-a"/><textField type="general" position="2" xmlns:f="urn:f" f:opaque="field-b"/></textFields></textPr><parameters count="2"><parameter name="A" xmlns:f="urn:f" f:opaque="parameter-a"/><parameter name="B" xmlns:f="urn:f" f:opaque="parameter-b"/></parameters></connection></connections>"#,
            MCE = litchi_ooxml_common::mce::NAMESPACE,
        );
        let before = Connections::parse(source.as_bytes()).unwrap();
        let mut after = before.clone();
        after.connections[0].web.as_mut().unwrap().tables = Some(vec![
            WebTableSelector::String("B".into()),
            WebTableSelector::Index(4),
        ]);
        after.connections[0].text.as_mut().unwrap().fields = Some(vec![
            TextField {
                field_type: Some(TextFieldType::General),
                position: Some(2),
            },
            TextField {
                field_type: Some(TextFieldType::Text),
                position: Some(1),
            },
        ]);
        after.connections[0].parameters = Some(vec![
            after.connections[0].parameters.as_ref().unwrap()[1].clone(),
            after.connections[0].parameters.as_ref().unwrap()[0].clone(),
        ]);

        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let text = std::str::from_utf8(&patched).unwrap();
        for marker in [
            "f:opaque=\"table-a\"",
            "f:opaque=\"table-b\"",
            "f:opaque=\"field-a\"",
            "f:opaque=\"field-b\"",
            "f:opaque=\"parameter-a\"",
            "f:opaque=\"parameter-b\"",
            "<mc:AlternateContent>",
            "<u:inactive/>",
        ] {
            assert!(text.contains(marker), "missing preserved markup {marker}");
        }
        assert!(text.find("<s v=\"B\"").unwrap() < text.find("<x v=\"4\"").unwrap());
        assert!(
            text.find("f:opaque=\"field-b\"").unwrap() < text.find("f:opaque=\"field-a\"").unwrap()
        );
        assert!(
            text.find("f:opaque=\"parameter-b\"").unwrap()
                < text.find("f:opaque=\"parameter-a\"").unwrap()
        );
        assert_eq!(Connections::parse(&patched).unwrap(), after);
    }

    #[test]
    fn source_patch_removes_selected_nested_subtrees_with_their_owned_markup() {
        let source = format!(
            r#"<connections xmlns="{CORE_NAMESPACE}" xmlns:mc="{MCE}" xmlns:u="urn:u"><connection id="1" refreshedVersion="7"><webPr><?before?><tables count="1"><s v="A"><?inside?><mc:AlternateContent><mc:Choice Requires="u"><u:inactive/></mc:Choice><mc:Fallback/></mc:AlternateContent></s></tables><?after?></webPr></connection></connections>"#,
            MCE = litchi_ooxml_common::mce::NAMESPACE,
        );
        let before = Connections::parse(source.as_bytes()).unwrap();
        let mut after = before.clone();
        after.connections[0].web.as_mut().unwrap().tables = None;

        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let text = std::str::from_utf8(&patched).unwrap();
        assert!(text.contains("<?before?>") && text.contains("<?after?>"));
        assert!(!text.contains("<tables") && !text.contains("<?inside?>"));
        assert_eq!(Connections::parse(&patched).unwrap(), after);
    }

    #[test]
    fn source_patch_publishes_connection_reorders() {
        let source = format!(
            r#"<connections xmlns="{CORE_NAMESPACE}"><connection id="1" refreshedVersion="7" name="one"/><connection id="2" refreshedVersion="7" name="two"/></connections>"#
        );
        let before = Connections::parse(source.as_bytes()).unwrap();
        let mut after = before.clone();
        after.connections.swap(0, 1);

        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let text = std::str::from_utf8(&patched).unwrap();
        assert!(text.find("id=\"2\"").unwrap() < text.find("id=\"1\"").unwrap());
        assert_eq!(Connections::parse(&patched).unwrap(), after);
    }

    #[test]
    fn source_edits_reject_oversized_final_output_before_copying() {
        let source = vec![b'x'; MAX_XML_BYTES];
        let error = apply_source_edits(
            &source,
            vec![SourceEdit {
                range: 0..0,
                replacement: vec![b'y'],
            }],
        )
        .unwrap_err();
        assert_eq!(error.to_string(), "connection source output exceeds 16 MiB");
    }

    #[test]
    fn source_patch_preserves_mce_fallback_and_keeps_equal_offset_inserts_ordered() {
        let source = format!(
            r#"<p:connections xmlns:p="{CORE_NAMESPACE}" xmlns:mc="{MCE}" xmlns:u="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="u"><u:connection id="9" refreshedVersion="7"/></mc:Choice><mc:Fallback><p:connection id="1" refreshedVersion="7" name="before"/></mc:Fallback></mc:AlternateContent></p:connections>"#,
            MCE = litchi_ooxml_common::mce::NAMESPACE,
        );
        let before = Connections::parse(source.as_bytes()).unwrap();
        let mut after = before.clone();
        after.connections[0].name = Some("after".into());
        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let text = std::str::from_utf8(&patched).unwrap();
        assert!(text.contains("mc:AlternateContent"));
        assert!(text.contains("mc:Fallback"));
        assert!(text.contains("u:connection"));
        assert_eq!(Connections::parse(&patched).unwrap(), after);

        let source = format!(
            r#"<p:connections xmlns:p="{CORE_NAMESPACE}"><p:connection id="1" refreshedVersion="7"/></p:connections>"#
        );
        let before = Connections::parse(source.as_bytes()).unwrap();
        let mut after = before.clone();
        after.connections.push(connection(2, None));
        after.connections.push(connection(3, None));
        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let actual = Connections::parse(&patched).unwrap();
        assert_eq!(
            actual
                .connections
                .iter()
                .map(|connection| connection.id)
                .collect::<Vec<_>>(),
            vec![1, 2, 3]
        );
    }

    #[test]
    fn source_patch_updates_typed_children_inside_active_mce_fallback() {
        let source = format!(
            r#"<p:connections xmlns:p="{CORE_NAMESPACE}" xmlns:mc="{MCE}" xmlns:u="urn:unsupported"><p:connection id="1" refreshedVersion="7"><mc:AlternateContent><mc:Choice Requires="u"><u:dbPr connection="inactive"/></mc:Choice><mc:Fallback><p:dbPr connection="before"/></mc:Fallback></mc:AlternateContent></p:connection></p:connections>"#,
            MCE = litchi_ooxml_common::mce::NAMESPACE,
        );
        let before = Connections::parse(source.as_bytes()).unwrap();
        let mut after = before.clone();
        after.connections[0].database = Some(database_connection());
        let patched = patch_connections_source(source.as_bytes(), &before, &after, false).unwrap();
        let text = std::str::from_utf8(&patched).unwrap();
        assert!(text.contains("connection=\"inactive\""));
        assert!(text.contains("connection=\"Provider=opaque\""));
        assert_eq!(Connections::parse(&patched).unwrap(), after);
    }

    #[test]
    fn opaque_namespace_rewrite_does_not_touch_text_or_attribute_values() {
        let source = format!(
            r#"<p:extLst xmlns:p="{CORE_NAMESPACE}" xmlns:f="{CORE_NAMESPACE}" marker="{CORE_NAMESPACE}"><p:ext><f:item>{CORE_NAMESPACE}</f:item></p:ext></p:extLst>"#
        );
        let rewritten = rewrite_namespace_declarations(source.as_bytes(), true).unwrap();
        let text = std::str::from_utf8(&rewritten).unwrap();
        assert!(text.contains(&format!("xmlns:p=\"{STRICT_NAMESPACE}\"")));
        assert!(text.contains(&format!("xmlns:f=\"{STRICT_NAMESPACE}\"")));
        assert!(text.contains(&format!("marker=\"{CORE_NAMESPACE}\"")));
        assert!(text.contains(&format!(">{CORE_NAMESPACE}</f:item>")));
    }
}

#[derive(Debug)]
struct SourceTree {
    nodes: Vec<SourceNode>,
    root: usize,
}

#[derive(Debug)]
struct SourceNode {
    qualified: String,
    namespace: NamespaceUri,
    local: String,
    context: NamespaceContext,
    parent: Option<usize>,
    start: usize,
    end_start: usize,
    end: usize,
    close_pos: usize,
    self_closing: bool,
    active: bool,
    attrs: Vec<SourceAttribute>,
    children: Vec<usize>,
}

#[derive(Debug)]
struct SourceAttribute {
    namespace: NamespaceUri,
    local: String,
    namespace_declaration: bool,
    start: usize,
    value_start: usize,
    value_end: usize,
}

struct SourceStartTag {
    qualified: String,
    namespace: NamespaceUri,
    local: String,
    context: NamespaceContext,
    attrs: Vec<SourceAttribute>,
    close_pos: usize,
    self_closing: bool,
}

fn is_mce_wrapper(node: &SourceNode) -> bool {
    node.namespace.as_ref() == litchi_ooxml_common::mce::NAMESPACE
        && matches!(
            node.local.as_str(),
            "AlternateContent" | "Choice" | "Fallback"
        )
}

impl SourceTree {
    fn parse(source: &[u8]) -> Result<Self> {
        let mut tree = Self::parse_raw_with_context(source, namespace_root())?;
        tree.mark_active(source)?;
        Ok(tree)
    }

    fn parse_with_bindings(source: &[u8], inherited: &[(String, String)]) -> Result<Self> {
        let mut tree = Self::parse_raw_with_context(source, namespace_from_bindings(inherited)?)?;
        tree.mark_active(source)?;
        Ok(tree)
    }

    fn parse_raw(source: &[u8]) -> Result<Self> {
        Self::parse_raw_with_context(source, namespace_root())
    }

    fn parse_raw_with_context(source: &[u8], inherited_context: NamespaceContext) -> Result<Self> {
        let mut nodes: Vec<SourceNode> = Vec::new();
        let mut stack: Vec<usize> = Vec::new();
        let mut root: Option<usize> = None;
        let mut position = 0;
        while position < source.len() {
            if source[position] != b'<' {
                position += 1;
                continue;
            }
            if source[position..].starts_with(b"<?") {
                position = find_source_bytes(source, position + 2, b"?>")? + 2;
                continue;
            }
            if source[position..].starts_with(b"<!--") {
                position = find_source_bytes(source, position + 4, b"-->")? + 3;
                continue;
            }
            if source[position..].starts_with(b"<![CDATA[") {
                position = find_source_bytes(source, position + 9, b"]]>")? + 3;
                continue;
            }
            if source[position..].starts_with(b"<!") {
                position = source_tag_end(source, position)? + 1;
                continue;
            }
            if source[position..].starts_with(b"</") {
                let end = source_tag_end(source, position)?;
                let name_start = position + 2;
                let name_end = source_name_end(source, name_start);
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("connection source has an unmatched closing tag"))?;
                let closing =
                    std::str::from_utf8(&source[name_start..name_end]).map_err(xml_error)?;
                if nodes[node].qualified != closing {
                    return Err(invalid("connection source has mismatched tags"));
                }
                nodes[node].end_start = position;
                nodes[node].end = end + 1;
                position = end + 1;
                continue;
            }
            let end = source_tag_end(source, position)?;
            if nodes.len() >= MAX_DOM_NODES {
                return Err(invalid("connection source node limit exceeded"));
            }
            if stack.len() >= MAX_DOM_DEPTH {
                return Err(invalid("connection source depth limit exceeded"));
            }
            nodes
                .try_reserve(1)
                .map_err(|_source| invalid("connection source node allocation failed"))?;
            let parent = stack.last().copied();
            let inherited = parent
                .and_then(|node| nodes.get(node))
                .map(|node| node.context.clone())
                .unwrap_or_else(|| inherited_context.clone());
            let SourceStartTag {
                qualified,
                namespace,
                local,
                context,
                attrs,
                close_pos,
                self_closing,
            } = source_start_tag(source, position, end, &inherited)?;
            let node = nodes.len();
            nodes.push(SourceNode {
                qualified,
                namespace,
                local,
                context,
                parent,
                start: position,
                end_start: end + 1,
                end: end + 1,
                close_pos,
                self_closing,
                active: true,
                attrs,
                children: Vec::new(),
            });
            if let Some(parent) = stack.last().copied() {
                nodes[parent]
                    .children
                    .try_reserve(1)
                    .map_err(|_source| invalid("connection source child allocation failed"))?;
                nodes[parent].children.push(node);
            } else if root.replace(node).is_some() {
                return Err(invalid("connection source has multiple roots"));
            }
            if !self_closing {
                stack
                    .try_reserve(1)
                    .map_err(|_source| invalid("connection source depth allocation failed"))?;
                stack.push(node);
            }
            position = end + 1;
        }
        if !stack.is_empty() {
            return Err(invalid("connection source has unterminated markup"));
        }
        Ok(Self {
            nodes,
            root: root.ok_or_else(|| invalid("connection source has no root"))?,
        })
    }

    fn mark_active(&mut self, source: &[u8]) -> Result<()> {
        let mut offsets = Vec::new();
        offsets
            .try_reserve_exact(self.nodes.len())
            .map_err(|_source| invalid("connection source offset allocation failed"))?;
        for node in &self.nodes {
            if !node.namespace.is_empty() {
                offsets.push(
                    u32::try_from(node.start).map_err(|_error| {
                        invalid("connection source offset exceeds 32-bit bounds")
                    })?,
                );
            }
        }
        if offsets.is_empty() {
            return Ok(());
        }
        let Some(active) = active_source_offsets(source, &offsets)? else {
            return Ok(());
        };
        let mut active_set = HashSet::new();
        active_set
            .try_reserve(active.len())
            .map_err(|_source| invalid("connection source active-set allocation failed"))?;
        active_set.extend(active);
        for node in &mut self.nodes {
            if !node.namespace.is_empty() {
                node.active =
                    active_set.contains(&u32::try_from(node.start).map_err(|_error| {
                        invalid("connection source offset exceeds 32-bit bounds")
                    })?);
            }
        }
        Ok(())
    }

    fn attribute(&self, source: &[u8], node: usize, name: &str) -> Result<Option<String>> {
        self.nodes[node]
            .attrs
            .iter()
            .find(|attribute| {
                !attribute.namespace_declaration
                    && attribute.namespace.is_empty()
                    && attribute.local == name
            })
            .map(|attribute| decode_source_attribute(source, attribute))
            .transpose()
    }

    fn child(&self, node: usize, name: &str) -> Result<Option<usize>> {
        let mut matches = Vec::new();
        let namespace = self.nodes[node].namespace.as_ref();
        self.collect_semantic_children(node, namespace, name, &mut matches);
        if matches.len() > 1 {
            return Err(invalid(format!("connection source has duplicate '{name}'")));
        }
        Ok(matches.into_iter().next())
    }

    fn collect_semantic_children(
        &self,
        node: usize,
        namespace: &str,
        name: &str,
        matches: &mut Vec<usize>,
    ) {
        for child in &self.nodes[node].children {
            let child = *child;
            let value = &self.nodes[child];
            if !value.active {
                continue;
            }
            if value.namespace.as_ref() == namespace && value.local == name {
                matches.push(child);
                continue;
            }
            if is_mce_wrapper(value) {
                self.collect_semantic_children(child, namespace, name, matches);
            }
        }
    }

    fn semantic_insertion_parent(&self, node: usize, name: &str) -> Option<usize> {
        let mut children = Vec::new();
        let namespace = self.nodes[node].namespace.as_ref();
        self.collect_semantic_elements(node, namespace, &mut children);
        let order = child_order(name);
        children
            .iter()
            .find(|child| child_order(&self.nodes[**child].local) > order)
            .or_else(|| children.last())
            .and_then(|child| self.nodes[*child].parent)
    }

    fn collect_semantic_elements(&self, node: usize, namespace: &str, output: &mut Vec<usize>) {
        for child in &self.nodes[node].children {
            let child = *child;
            let value = &self.nodes[child];
            if !value.active {
                continue;
            }
            if value.namespace.as_ref() == namespace {
                output.push(child);
            } else if is_mce_wrapper(value) {
                self.collect_semantic_elements(child, namespace, output);
            }
        }
    }

    fn semantic_elements(&self, node: usize) -> Result<Vec<usize>> {
        let namespace = self.nodes[node].namespace.as_ref();
        let mut elements = Vec::new();
        elements
            .try_reserve(self.nodes.len().min(MAX_DOM_NODES))
            .map_err(|_source| invalid("connection source child allocation failed"))?;
        self.collect_semantic_elements(node, namespace, &mut elements);
        Ok(elements)
    }

    fn connection_nodes(&self) -> Result<Vec<usize>> {
        let namespace = self.nodes[self.root].namespace.as_ref();
        let mut connections = Vec::new();
        // The caller validates the count against MAX_CONNECTIONS; reserve the
        // bounded source-node count so this metadata does not grow by panic-on-OOM.
        connections
            .try_reserve(self.nodes.len().min(MAX_CONNECTIONS))
            .map_err(|_source| invalid("connection source node allocation failed"))?;
        for (index, node) in self.nodes.iter().enumerate() {
            if node.active
                && node.namespace.as_ref() == namespace
                && node.local == "connection"
                && self.is_semantic_connection_child(index)
            {
                if connections.len() >= MAX_CONNECTIONS {
                    return Err(invalid("connection source connection limit exceeded"));
                }
                connections.push(index);
            }
        }
        Ok(connections)
    }

    fn is_semantic_connection_child(&self, node: usize) -> bool {
        let mut parent = self.nodes[node].parent;
        while let Some(index) = parent {
            if index == self.root {
                return true;
            }
            let value = &self.nodes[index];
            if !is_mce_wrapper(value) {
                return false;
            }
            parent = value.parent;
        }
        false
    }
}

fn find_source_bytes(source: &[u8], start: usize, needle: &[u8]) -> Result<usize> {
    source[start..]
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|position| start + position)
        .ok_or_else(|| invalid("connection source has an unterminated declaration"))
}

const MCE_NAMESPACE_MARKER: &[u8] =
    b"<!--http://schemas.openxmlformats.org/markup-compatibility/2006-->";

fn active_source_offsets(source: &[u8], offsets: &[u32]) -> Result<Option<Vec<u32>>> {
    if offsets.is_empty() {
        return Ok(Some(Vec::new()));
    }
    let raw_mce_namespace = source
        .windows(litchi_ooxml_common::mce::NAMESPACE.len())
        .any(|window| window == litchi_ooxml_common::mce::NAMESPACE.as_bytes());
    let decoded_mce_namespace = if raw_mce_namespace || source.contains(&b'&') {
        has_mce_namespace_declaration(source)?
    } else {
        false
    };
    if !decoded_mce_namespace {
        return Ok(Some(offsets.to_vec()));
    }
    if source.len() > MAX_XML_BYTES {
        return Err(invalid("connection source exceeds 16 MiB"));
    }
    let without_pi_source = strip_processing_instructions(source)?;
    let raw_mce_namespace = without_pi_source
        .as_ref()
        .windows(litchi_ooxml_common::mce::NAMESPACE.len())
        .any(|window| window == litchi_ooxml_common::mce::NAMESPACE.as_bytes());
    let mce_source = if raw_mce_namespace {
        Cow::Borrowed(source)
    } else {
        let capacity = source
            .len()
            .checked_add(MCE_NAMESPACE_MARKER.len())
            .ok_or_else(|| invalid("connection source MCE marker size overflowed"))?;
        let mut marked = Vec::new();
        marked
            .try_reserve_exact(capacity)
            .map_err(|_| invalid("connection source MCE marker allocation failed"))?;
        marked.extend_from_slice(source);
        marked.extend_from_slice(MCE_NAMESPACE_MARKER);
        Cow::Owned(marked)
    };
    let without_pi = if raw_mce_namespace {
        without_pi_source
    } else {
        strip_processing_instructions(mce_source.as_ref())?
    };
    let max_marked_bytes = MAX_XML_BYTES.saturating_mul(4);
    let processing = litchi_ooxml_common::mce::Limits {
        max_input_bytes: max_marked_bytes,
        ..Default::default()
    };
    let limits = litchi_ooxml_common::mce::OffsetLimits {
        max_source_bytes: mce_source.len(),
        max_marked_bytes,
        processing,
        ..Default::default()
    };
    let capabilities = litchi_ooxml_common::mce::Capabilities::default();
    if matches!(&without_pi, Cow::Borrowed(_)) {
        return litchi_ooxml_common::mce::active_offsets(
            mce_source.as_ref(),
            offsets,
            &capabilities,
            &limits,
        )
        .map(Some)
        .map_err(Into::into);
    }

    let mut mapped = Vec::new();
    mapped
        .try_reserve_exact(offsets.len())
        .map_err(|_source| invalid("connection source offset mapping allocation failed"))?;
    mapped.extend_from_slice(offsets);
    mapped.sort_unstable();
    let mut originals = Vec::new();
    originals
        .try_reserve_exact(mapped.len())
        .map_err(|_source| invalid("connection source offset mapping allocation failed"))?;
    originals.extend_from_slice(&mapped);
    processing_instruction_ranges(source, &mut mapped)?;
    let selected = litchi_ooxml_common::mce::active_offsets(
        without_pi.as_ref(),
        &mapped,
        &capabilities,
        &limits,
    )
    .map_err(xml_error)?;
    let selected = selected.into_iter().collect::<HashSet<_>>();
    let selected_original = originals
        .into_iter()
        .zip(mapped)
        .filter_map(|(original, mapped)| selected.contains(&mapped).then_some(original))
        .collect::<HashSet<_>>();
    Ok(Some(
        offsets
            .iter()
            .copied()
            .filter(|offset| selected_original.contains(offset))
            .collect(),
    ))
}

fn has_mce_namespace_declaration(source: &[u8]) -> Result<bool> {
    let mut reader = Reader::from_reader(source);
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element)) => {
                let decoder = reader.decoder();
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(xml_error)?;
                    let key = attribute.key.as_ref();
                    if key != b"xmlns" && !key.starts_with(b"xmlns:") {
                        continue;
                    }
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                        .map_err(xml_error)?;
                    if value.as_ref() == litchi_ooxml_common::mce::NAMESPACE {
                        return Ok(true);
                    }
                }
            },
            Ok(Event::Eof) => return Ok(false),
            Ok(_) => {},
            Err(error) => return Err(xml_error(error)),
        }
    }
}

fn processing_instruction_ranges(source: &[u8], offsets: &mut [u32]) -> Result<()> {
    let mut reader = Reader::from_reader(source);
    // `offsets` are byte offsets into `source`, whose leading byte-order mark
    // precedes reader position zero.
    let origin = ReaderOrigin::of(source);
    let position = |reader: &Reader<&[u8]>| {
        origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("processing-instruction source offset exceeds usize"))
    };
    let mut offset_index = 0usize;
    let mut removed = 0usize;
    loop {
        let start = position(&reader)?;
        match reader.read_event() {
            Ok(Event::PI(_)) => {
                let end = position(&reader)?;
                if start > end || end > source.len() {
                    return Err(invalid("invalid processing-instruction source span"));
                }
                while offset_index < offsets.len() {
                    let offset = usize::try_from(offsets[offset_index]).map_err(|_error| {
                        invalid("connection source offset exceeds platform bounds")
                    })?;
                    if offset >= start {
                        break;
                    }
                    offsets[offset_index] =
                        u32::try_from(offset.checked_sub(removed).ok_or_else(|| {
                            invalid("connection source offset mapping underflowed")
                        })?)
                        .map_err(|_error| {
                            invalid("connection source offset exceeds 32-bit bounds")
                        })?;
                    offset_index += 1;
                }
                if offset_index < offsets.len() {
                    let offset = usize::try_from(offsets[offset_index]).map_err(|_error| {
                        invalid("connection source offset exceeds platform bounds")
                    })?;
                    if offset < end {
                        return Err(invalid(
                            "connection source offset falls inside a processing instruction",
                        ));
                    }
                }
                removed = removed
                    .checked_add(end - start)
                    .ok_or_else(|| invalid("connection source offset mapping overflowed"))?;
            },
            Ok(Event::Eof) => {
                while offset_index < offsets.len() {
                    let offset = usize::try_from(offsets[offset_index]).map_err(|_error| {
                        invalid("connection source offset exceeds platform bounds")
                    })?;
                    offsets[offset_index] =
                        u32::try_from(offset.checked_sub(removed).ok_or_else(|| {
                            invalid("connection source offset mapping underflowed")
                        })?)
                        .map_err(|_error| {
                            invalid("connection source offset exceeds 32-bit bounds")
                        })?;
                    offset_index += 1;
                }
                break;
            },
            Ok(_) => {},
            Err(error) => return Err(xml_error(error)),
        }
    }
    Ok(())
}

fn source_tag_end(source: &[u8], start: usize) -> Result<usize> {
    let mut quote = None;
    for (offset, byte) in source[start + 1..].iter().enumerate() {
        match (quote, byte) {
            (Some(value), byte) if *byte == value => quote = None,
            (None, b'\'' | b'\"') => quote = Some(*byte),
            (None, b'>') => return Ok(start + 1 + offset),
            _ => {},
        }
    }
    Err(invalid("connection source has an unterminated tag"))
}

fn source_start_tag(
    source: &[u8],
    start: usize,
    end: usize,
    inherited: &NamespaceContext,
) -> Result<SourceStartTag> {
    let name_start = start + 1;
    let name_end = source_name_end(source, name_start);
    if name_start == name_end {
        return Err(invalid("connection source has an empty element name"));
    }
    let qualified = std::str::from_utf8(&source[name_start..name_end])
        .map_err(xml_error)?
        .to_owned();
    let mut raw_attributes = Vec::new();
    let mut position = name_end;
    while position < end {
        while position < end && source[position].is_ascii_whitespace() {
            position += 1;
        }
        if position >= end || source[position] == b'/' {
            break;
        }
        let attr_start = position;
        let attr_end = source_name_end(source, position);
        if attr_end == attr_start {
            return Err(invalid("connection source has an invalid attribute"));
        }
        position = attr_end;
        while position < end && source[position].is_ascii_whitespace() {
            position += 1;
        }
        if position >= end || source[position] != b'=' {
            return Err(invalid("connection source attribute is missing '='"));
        }
        position += 1;
        while position < end && source[position].is_ascii_whitespace() {
            position += 1;
        }
        let quote = *source
            .get(position)
            .ok_or_else(|| invalid("connection source attribute is missing quotes"))?;
        if !matches!(quote, b'\'' | b'\"') {
            return Err(invalid("connection source attribute is missing quotes"));
        }
        position += 1;
        let value_start = position;
        while position < end && source[position] != quote {
            position += 1;
        }
        if position >= end {
            return Err(invalid("connection source attribute is unterminated"));
        }
        let value_end = position;
        position += 1;
        let qualified_attribute = std::str::from_utf8(&source[attr_start..attr_end])
            .map_err(xml_error)?
            .to_owned();
        raw_attributes
            .try_reserve(1)
            .map_err(|_source| invalid("connection source attribute allocation failed"))?;
        raw_attributes.push((qualified_attribute, attr_start, value_start, value_end));
    }
    let mut decoded_declarations = Vec::new();
    decoded_declarations
        .try_reserve(raw_attributes.len())
        .map_err(|_source| invalid("connection source namespace allocation failed"))?;
    for (qualified_attribute, _start, value_start, value_end) in &raw_attributes {
        let (prefix, local) = source_name_parts(qualified_attribute)?;
        if qualified_attribute == "xmlns" || prefix == "xmlns" {
            let binding = if qualified_attribute == "xmlns" {
                ""
            } else {
                local
            };
            let value = decode_source_attribute_value(source, *value_start, *value_end)?;
            decoded_declarations.push((binding.to_owned(), value));
        }
    }
    let declaration_refs = decoded_declarations
        .iter()
        .map(|(prefix, uri)| NamespaceDecl::new(prefix, uri));
    let context = inherited.child(declaration_refs)?;
    let mut attributes = Vec::new();
    attributes
        .try_reserve(raw_attributes.len())
        .map_err(|_source| invalid("connection source attribute allocation failed"))?;
    for (qualified_attribute, attr_start, value_start, value_end) in raw_attributes {
        let (prefix, local) = source_name_parts(&qualified_attribute)?;
        let namespace_declaration = qualified_attribute == "xmlns" || prefix == "xmlns";
        attributes.push(SourceAttribute {
            namespace: if namespace_declaration {
                Arc::clone(context.xmlns_uri())
            } else if prefix.is_empty() {
                Arc::from("")
            } else {
                resolve_source_binding(&context, prefix)?
            },
            local: local.to_owned(),
            namespace_declaration,
            start: attr_start,
            value_start,
            value_end,
        });
    }
    let self_closing = source[..end]
        .iter()
        .rposition(|byte| !byte.is_ascii_whitespace())
        .is_some_and(|position| source[position] == b'/');
    let close_pos = if self_closing {
        source[..end]
            .iter()
            .rposition(|byte| !byte.is_ascii_whitespace())
            .unwrap_or(end)
    } else {
        end
    };
    let (prefix, local) = source_name_parts(&qualified)?;
    let prefix = prefix.to_owned();
    let local = local.to_owned();
    let namespace = if prefix.is_empty() {
        resolve_source_binding(&context, "")?
    } else {
        resolve_source_binding(&context, &prefix)?
    };
    Ok(SourceStartTag {
        qualified,
        namespace,
        local,
        context,
        attrs: attributes,
        close_pos,
        self_closing,
    })
}

fn source_name_end(source: &[u8], mut position: usize) -> usize {
    while position < source.len()
        && !source[position].is_ascii_whitespace()
        && !matches!(source[position], b'/' | b'>' | b'=')
    {
        position += 1;
    }
    position
}

fn source_name_parts(value: &str) -> Result<(&str, &str)> {
    match value.split_once(':') {
        Some((prefix, local))
            if !prefix.is_empty() && !local.is_empty() && !local.contains(':') =>
        {
            Ok((prefix, local))
        },
        Some(_) => Err(invalid("connection source has an invalid QName")),
        None if !value.is_empty() => Ok(("", value)),
        None => Err(invalid("connection source has an empty QName")),
    }
}

fn resolve_source_binding(context: &NamespaceContext, prefix: &str) -> Result<NamespaceUri> {
    context.resolve_uri(prefix).cloned().ok_or_else(|| {
        invalid(format!(
            "connection source has an unbound prefix '{prefix}'"
        ))
    })
}

fn decode_source_attribute_value(
    source: &[u8],
    value_start: usize,
    value_end: usize,
) -> Result<String> {
    let value = std::str::from_utf8(
        source
            .get(value_start..value_end)
            .ok_or_else(|| invalid("connection source attribute range is invalid"))?,
    )
    .map_err(xml_error)?;
    quick_xml::escape::unescape(value)
        .map(|value| value.into_owned())
        .map_err(xml_error)
}

fn decode_source_attribute(source: &[u8], attribute: &SourceAttribute) -> Result<String> {
    decode_source_attribute_value(source, attribute.value_start, attribute.value_end)
}

#[cfg(test)]
mod source_offsets_tests {
    include!("codec/source_offsets_tests.rs");
}
