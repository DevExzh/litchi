//! Bounded RDF/XML parsing and deterministic serialization.

use super::{Graph, Object, Subject, Triple};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

pub(super) const RDF: &str = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
const RDF_NS: &[u8] = b"http://www.w3.org/1999/02/22-rdf-syntax-ns#";
const XML_NS: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const MAX_XML: usize = 16 * 1024 * 1024;
pub(super) const MAX_TRIPLES: usize = 65_536;
const MAX_VALUE: usize = 1024 * 1024;

#[derive(Clone)]
pub(super) struct TripleSpan {
    pub(super) start: usize,
    pub(super) end: usize,
    pub(super) subject: Subject,
}

pub(super) struct GraphParse {
    pub(super) graph: Graph,
    pub(super) spans: Vec<TripleSpan>,
    pub(super) root_close: usize,
}

pub(super) fn parse(path: &str, xml: &str) -> Result<GraphParse> {
    if xml.len() > MAX_XML {
        return invalid("RDF/XML part exceeds size limit");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root_depth = None;
    let mut root_close = None;
    let mut prefixes = Vec::new();
    let mut base = None;
    let mut subject: Option<(usize, Subject)> = None;
    let mut property: Option<(usize, usize, String, Object)> = None;
    let mut triples = Vec::new();
    let mut spans = Vec::new();
    let mut skip_depth = None;
    loop {
        let start = position(&reader)?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid RDF/XML in '{path}': {error}"))
            })?;
        let ns = resolved_namespace(&resolved);
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                if root_depth.is_none() {
                    if ns.as_deref() != Some(RDF_NS) || element.local_name().as_ref() != b"RDF" {
                        return invalid("RDF/XML root must be rdf:RDF");
                    }
                    root_depth = Some(depth);
                    for raw_attribute in element.attributes().with_checks(true) {
                        let attribute = raw_attribute.map_err(|error| {
                            Error::InvalidFormat(format!("invalid RDF root attribute: {error}"))
                        })?;
                        let key = std::str::from_utf8(attribute.key.as_ref()).map_err(|error| {
                            Error::InvalidFormat(format!(
                                "invalid RDF root attribute name: {error}"
                            ))
                        })?;
                        let value = decode_attr(&reader, &attribute)?;
                        if key == "xml:base" {
                            base = Some(value);
                        } else if key == "xmlns" {
                            prefixes.push((String::new(), value));
                        } else if let Some(prefix) = key.strip_prefix("xmlns:") {
                            prefixes.push((prefix.to_string(), value));
                        }
                    }
                } else if skip_depth.is_none()
                    && subject.is_none()
                    && root_depth.is_some_and(|root| depth == root + 1)
                {
                    if ns.as_deref() == Some(RDF_NS)
                        && element.local_name().as_ref() == b"Description"
                    {
                        subject = Some((depth, parse_subject(&reader, &element)?));
                    } else {
                        skip_depth = Some(depth);
                    }
                } else if skip_depth.is_none()
                    && property.is_none()
                    && subject.as_ref().is_some_and(|item| depth == item.0 + 1)
                {
                    let predicate = predicate_iri(ns.as_deref(), element.local_name().as_ref())?;
                    let object =
                        parse_object_attrs(&reader, &element)?.unwrap_or(Object::Literal {
                            value: String::new(),
                            datatype: None,
                            language: None,
                        });
                    property = Some((depth, start, predicate, object));
                } else if skip_depth.is_none() && property.is_some() {
                    return invalid(
                        "nested RDF property content is not supported for typed mutation",
                    );
                }
                depth += 1;
            },
            Event::Empty(element) => {
                if subject.is_none()
                    && root_depth.is_some_and(|root| depth == root + 1)
                    && ns.as_deref() == Some(RDF_NS)
                    && element.local_name().as_ref() == b"Description"
                {
                    let _ = parse_subject(&reader, &element)?;
                } else if let Some((subject_depth, current)) = &subject
                    && depth == *subject_depth + 1
                {
                    let predicate = predicate_iri(ns.as_deref(), element.local_name().as_ref())?;
                    let object =
                        parse_object_attrs(&reader, &element)?.unwrap_or(Object::Literal {
                            value: String::new(),
                            datatype: None,
                            language: None,
                        });
                    triples.push(Triple {
                        subject: current.clone(),
                        predicate,
                        object,
                    });
                    spans.push(TripleSpan {
                        start,
                        end,
                        subject: current.clone(),
                    });
                }
            },
            Event::Text(text_event) if property.is_some() => {
                let text = text_event
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid RDF literal: {error}"))
                    })?;
                if let Some((
                    _,
                    _,
                    _,
                    Object::Literal {
                        value: literal_value,
                        ..
                    },
                )) = property.as_mut()
                {
                    literal_value.push_str(&text);
                    if literal_value.len() > MAX_VALUE {
                        return invalid("RDF literal exceeds size limit");
                    }
                } else if !text.trim().is_empty() {
                    return invalid("RDF resource property cannot contain text");
                }
            },
            Event::CData(cdata_event) if property.is_some() => {
                let text = cdata_event
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid RDF literal: {error}"))
                    })?;
                if let Some((
                    _,
                    _,
                    _,
                    Object::Literal {
                        value: literal_value,
                        ..
                    },
                )) = property.as_mut()
                {
                    literal_value.push_str(&text);
                    if literal_value.len() > MAX_VALUE {
                        return invalid("RDF literal exceeds size limit");
                    }
                } else if !text.trim().is_empty() {
                    return invalid("RDF resource property cannot contain CDATA");
                }
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("RDF/XML depth underflow".to_string()))?;
                if skip_depth == Some(depth) {
                    skip_depth = None;
                } else if property.as_ref().is_some_and(|item| item.0 == depth) {
                    let (_, property_start, predicate, object) =
                        property.take().ok_or_else(|| {
                            Error::InvalidFormat(
                                "RDF property closed without an active property".to_string(),
                            )
                        })?;
                    let current = subject
                        .as_ref()
                        .ok_or_else(|| {
                            Error::InvalidFormat("RDF property has no active subject".to_string())
                        })?
                        .1
                        .clone();
                    triples.push(Triple {
                        subject: current.clone(),
                        predicate,
                        object,
                    });
                    spans.push(TripleSpan {
                        start: property_start,
                        end,
                        subject: current,
                    });
                } else if subject.as_ref().is_some_and(|item| item.0 == depth) {
                    subject = None;
                } else if root_depth == Some(depth)
                    && ns.as_deref() == Some(RDF_NS)
                    && element.local_name().as_ref() == b"RDF"
                {
                    root_close = Some(start);
                }
            },
            Event::DocType(_) | Event::GeneralRef(_) => {
                return invalid("DTDs and entity references are prohibited in RDF/XML");
            },
            Event::PI(_) => return invalid("processing instructions are prohibited in RDF/XML"),
            Event::Eof => break,
            Event::Text(_) | Event::CData(_) | Event::Comment(_) | Event::Decl(_) => {},
        }
        if triples.len() > MAX_TRIPLES {
            return invalid("RDF graph exceeds triple limit");
        }
        buffer.clear();
    }
    let root_close_offset =
        root_close.ok_or_else(|| Error::InvalidFormat("unterminated rdf:RDF root".to_string()))?;
    Ok(GraphParse {
        graph: Graph {
            path: path.to_string(),
            base,
            prefixes,
            triples,
        },
        spans,
        root_close: root_close_offset,
    })
}

fn parse_subject(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<Subject> {
    let about = rdf_attr(reader, element, b"about")?;
    let node = rdf_attr(reader, element, b"nodeID")?;
    match (about, node) {
        (Some(value), None) => Ok(Subject::Iri(value)),
        (None, Some(value)) => Ok(Subject::BlankNode(value)),
        _ => invalid("rdf:Description requires exactly one of rdf:about or rdf:nodeID"),
    }
}

fn parse_object_attrs(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<Option<Object>> {
    let resource = rdf_attr(reader, element, b"resource")?;
    let node = rdf_attr(reader, element, b"nodeID")?;
    let datatype = rdf_attr(reader, element, b"datatype")?;
    let language = namespaced_attr(reader, element, XML_NS, b"lang")?;
    match (resource, node) {
        (Some(value), None) if datatype.is_none() && language.is_none() => {
            Ok(Some(Object::Iri(value)))
        },
        (None, Some(value)) if datatype.is_none() && language.is_none() => {
            Ok(Some(Object::BlankNode(value)))
        },
        (None, None) => Ok(Some(Object::Literal {
            value: String::new(),
            datatype,
            language,
        })),
        _ => invalid("invalid or ambiguous RDF property object attributes"),
    }
}

pub(super) fn serialize_graph(triples: &[Triple]) -> Result<String> {
    let mut out =
        format!("<?xml version=\"1.0\" encoding=\"UTF-8\"?><rdf:RDF xmlns:rdf=\"{RDF}\">");
    for triple in triples {
        out.push_str(&description_xml(triple)?);
    }
    out.push_str("</rdf:RDF>");
    Ok(out)
}

pub(super) fn description_xml(triple: &Triple) -> Result<String> {
    let mut out = "<rdf:Description".to_string();
    match &triple.subject {
        Subject::Iri(value) => attr(&mut out, "rdf:about", value)?,
        Subject::BlankNode(value) => attr(&mut out, "rdf:nodeID", value)?,
    }
    out.push('>');
    out.push_str(&property_xml(triple)?);
    out.push_str("</rdf:Description>");
    Ok(out)
}

pub(super) fn property_xml(triple: &Triple) -> Result<String> {
    let (namespace, local) = split_predicate(&triple.predicate)?;
    let mut out = format!("<p:{local} xmlns:p=\"{}\"", escape(&namespace));
    match &triple.object {
        Object::Iri(value) => {
            attr(&mut out, "rdf:resource", value)?;
            out.push_str("/>");
        },
        Object::BlankNode(value) => {
            attr(&mut out, "rdf:nodeID", value)?;
            out.push_str("/>");
        },
        Object::Literal {
            value,
            datatype,
            language,
        } => {
            if let Some(datatype_iri) = datatype {
                attr(&mut out, "rdf:datatype", datatype_iri)?;
            }
            if let Some(language_tag) = language {
                attr(&mut out, "xml:lang", language_tag)?;
            }
            out.push('>');
            out.push_str(&escape(value));
            out.push_str("</p:");
            out.push_str(local);
            out.push('>');
        },
    }
    Ok(out)
}

pub(super) fn validate_blank(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 256
        || !value
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        invalid("invalid RDF blank-node identifier")
    } else {
        Ok(())
    }
}

pub(super) fn validate_iri(value: &str) -> Result<()> {
    validate_value(value)?;
    if value.chars().any(char::is_whitespace) {
        invalid("RDF IRI contains whitespace")
    } else {
        Ok(())
    }
}

pub(super) fn validate_value(value: &str) -> Result<()> {
    if value.len() > MAX_VALUE
        || value
            .chars()
            .any(|ch| matches!(ch as u32, 0..=8 | 11 | 12 | 14..=31 | 0xFFFE | 0xFFFF))
    {
        invalid("RDF value exceeds limits or contains XML-prohibited characters")
    } else {
        Ok(())
    }
}

pub(super) fn external_iri(value: &str) -> bool {
    value.find(':').is_some_and(|index| {
        index > 0
            && value[..index]
                .bytes()
                .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'-' | b'.'))
    })
}

fn split_predicate(value: &str) -> Result<(String, &str)> {
    validate_iri(value)?;
    let index = value.rfind(['#', '/']).ok_or_else(|| {
        Error::InvalidFormat("RDF predicate IRI has no namespace boundary".to_string())
    })?;
    let local = &value[index + 1..];
    if local.is_empty()
        || !local.as_bytes()[0].is_ascii_alphabetic()
        || !local
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        return invalid("RDF predicate local name is not an XML name");
    }
    Ok((value[..=index].to_string(), local))
}

fn predicate_iri(namespace_bytes: Option<&[u8]>, local_bytes: &[u8]) -> Result<String> {
    let Some(bound_namespace) = namespace_bytes else {
        return invalid("RDF property requires a namespace");
    };
    let namespace_text = std::str::from_utf8(bound_namespace).map_err(|error| {
        Error::InvalidFormat(format!("invalid RDF predicate namespace: {error}"))
    })?;
    let local_name = std::str::from_utf8(local_bytes)
        .map_err(|error| Error::InvalidFormat(format!("invalid RDF predicate name: {error}")))?;
    Ok(format!("{namespace_text}{local_name}"))
}

fn rdf_attr(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    local: &[u8],
) -> Result<Option<String>> {
    namespaced_attr(reader, element, RDF_NS, local)
}

fn namespaced_attr(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    let mut result = None;
    for raw_attribute in element.attributes().with_checks(true) {
        let attribute = raw_attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid RDF attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(resolved, ResolveResult::Bound(Namespace(value)) if value == namespace)
            && name.as_ref() == local
        {
            if result.is_some() {
                return invalid("duplicate expanded RDF attribute");
            }
            result = Some(decode_attr(reader, &attribute)?);
        }
    }
    Ok(result)
}

pub(super) fn decode_attr(
    reader: &NsReader<&[u8]>,
    attr: &quick_xml::events::attributes::Attribute<'_>,
) -> Result<String> {
    attr.decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
        .map(std::borrow::Cow::into_owned)
        .map_err(|error| Error::InvalidFormat(format!("invalid RDF attribute value: {error}")))
}

fn resolved_namespace(resolution: &ResolveResult<'_>) -> Option<Vec<u8>> {
    match resolution {
        ResolveResult::Bound(Namespace(namespace)) => Some(namespace.to_vec()),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
    }
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|error| Error::InvalidFormat(format!("RDF XML position overflow: {error}")))
}

fn attr(out: &mut String, name: &str, value: &str) -> Result<()> {
    validate_value(value)?;
    out.push(' ');
    out.push_str(name);
    out.push_str("=\"");
    out.push_str(&escape(value));
    out.push('"');
    Ok(())
}

fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
}

pub(super) fn bounds(index: usize, len: usize) -> Error {
    Error::InvalidFormat(format!(
        "RDF triple index {index} is out of bounds for {len} triples"
    ))
}

pub(super) fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
