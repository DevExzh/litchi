//! Inert, bounded RDF/XML package metadata discovery and mutation.

use crate::core::OwnedPackage;
use crate::embedded_chart::{Addition, rebuild_package, splice};
use crate::{constants, media};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

const RDF: &str = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
const RDF_NS: &[u8] = b"http://www.w3.org/1999/02/22-rdf-syntax-ns#";
const XML_NS: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const MAX_XML: usize = 16 * 1024 * 1024;
const MAX_TRIPLES: usize = 65_536;
const MAX_VALUE: usize = 1024 * 1024;

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OdfRdfSubject {
    Iri(String),
    BlankNode(String),
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OdfRdfObject {
    Iri(String),
    BlankNode(String),
    Literal {
        value: String,
        datatype: Option<String>,
        language: Option<String>,
    },
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct OdfRdfTriple {
    pub subject: OdfRdfSubject,
    pub predicate: String,
    pub object: OdfRdfObject,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OdfRdfGraph {
    pub path: String,
    pub base: Option<String>,
    pub prefixes: Vec<(String, String)>,
    pub triples: Vec<OdfRdfTriple>,
}

#[derive(Clone)]
struct TripleSpan {
    start: usize,
    end: usize,
    subject: OdfRdfSubject,
}
struct Parsed {
    graph: OdfRdfGraph,
    spans: Vec<TripleSpan>,
    root_close: usize,
}

pub(crate) fn graphs(package: &OwnedPackage) -> Result<Vec<OdfRdfGraph>> {
    let archive = package.package()?;
    let mut paths: Vec<String> = archive
        .manifest()
        .entries
        .values()
        .filter(|entry| {
            entry.media_type == constants::ODF_MANIFEST_RDF_TYPE && !entry.full_path.ends_with('/')
        })
        .map(|entry| entry.full_path.clone())
        .collect();
    paths.sort();
    let mut result = Vec::with_capacity(paths.len());
    for path in paths {
        if !archive.has_file(&path) {
            return invalid(format!("RDF manifest entry '{path}' is dangling"));
        }
        let xml = String::from_utf8(archive.get_file(&path)?).map_err(|_| {
            Error::InvalidFormat(format!("RDF metadata part '{path}' is not UTF-8"))
        })?;
        result.push(parse(&path, &xml)?.graph);
    }
    Ok(result)
}

pub(crate) fn add_graph(
    package: &OwnedPackage,
    preferred: Option<&str>,
    triples: &[OdfRdfTriple],
) -> Result<(Vec<u8>, String)> {
    let path = match preferred {
        Some(path) => {
            let path = safe_path(path)?;
            if package.has_file(&path)? {
                return invalid(format!("RDF metadata path '{path}' already exists"));
            }
            path
        },
        None => unused_path(package)?,
    };
    validate_triples(package, triples, Some(&path))?;
    let xml = serialize_graph(triples)?;
    let content = String::from_utf8(package.get_file(constants::ODF_CONTENT)?)
        .map_err(|_| Error::InvalidFormat("content.xml is not UTF-8".to_string()))?;
    let bytes = rebuild_package(
        package,
        &content,
        vec![Addition {
            path: path.clone(),
            bytes: xml.into_bytes(),
            media_type: constants::ODF_MANIFEST_RDF_TYPE.to_string(),
        }],
        Vec::new(),
        Vec::new(),
        Vec::new(),
    )?;
    Ok((bytes, path))
}

pub(crate) fn replace_graph(
    package: &OwnedPackage,
    path: &str,
    triples: &[OdfRdfTriple],
) -> Result<Vec<u8>> {
    let path = existing_graph(package, path)?;
    validate_triples(package, triples, Some(&path))?;
    write_graph(package, &path, serialize_graph(triples)?)
}

pub(crate) fn remove_graph(package: &OwnedPackage, path: &str) -> Result<Vec<u8>> {
    let path = existing_graph(package, path)?;
    for graph in graphs(package)? {
        if graph.path != path
            && graph
                .triples
                .iter()
                .any(|triple| triple_refers_to(triple, &path))
        {
            return invalid(format!(
                "RDF metadata part '{path}' is still referenced by '{}'",
                graph.path
            ));
        }
    }
    let content = String::from_utf8(package.get_file(constants::ODF_CONTENT)?)
        .map_err(|_| Error::InvalidFormat("content.xml is not UTF-8".to_string()))?;
    rebuild_package(
        package,
        &content,
        Vec::new(),
        Vec::new(),
        vec![path],
        Vec::new(),
    )
}

pub(crate) fn add_triple(
    package: &OwnedPackage,
    path: &str,
    triple: &OdfRdfTriple,
) -> Result<(Vec<u8>, usize)> {
    let path = existing_graph(package, path)?;
    validate_triples(package, std::slice::from_ref(triple), Some(&path))?;
    let xml = graph_xml(package, &path)?;
    let parsed = parse(&path, &xml)?;
    if parsed.graph.triples.len() >= MAX_TRIPLES {
        return invalid("RDF graph exceeds triple limit");
    }
    let fragment = description_xml(triple)?;
    let updated = splice(&xml, parsed.root_close, parsed.root_close, &fragment)?;
    write_graph(package, &path, updated).map(|bytes| (bytes, parsed.graph.triples.len()))
}

pub(crate) fn replace_triple(
    package: &OwnedPackage,
    path: &str,
    index: usize,
    triple: &OdfRdfTriple,
) -> Result<Vec<u8>> {
    let path = existing_graph(package, path)?;
    validate_triples(package, std::slice::from_ref(triple), Some(&path))?;
    let xml = graph_xml(package, &path)?;
    let parsed = parse(&path, &xml)?;
    let span = parsed
        .spans
        .get(index)
        .ok_or_else(|| bounds(index, parsed.spans.len()))?;
    if triple.subject != span.subject {
        return invalid("replacing an RDF property cannot change its description subject");
    }
    let updated = splice(&xml, span.start, span.end, &property_xml(triple)?)?;
    write_graph(package, &path, updated)
}

pub(crate) fn remove_triple(package: &OwnedPackage, path: &str, index: usize) -> Result<Vec<u8>> {
    let path = existing_graph(package, path)?;
    let xml = graph_xml(package, &path)?;
    let parsed = parse(&path, &xml)?;
    let span = parsed
        .spans
        .get(index)
        .ok_or_else(|| bounds(index, parsed.spans.len()))?;
    write_graph(package, &path, splice(&xml, span.start, span.end, "")?)
}

pub(crate) fn move_triple(
    package: &OwnedPackage,
    path: &str,
    from: usize,
    to: usize,
) -> Result<Vec<u8>> {
    let path = existing_graph(package, path)?;
    let xml = graph_xml(package, &path)?;
    let parsed = parse(&path, &xml)?;
    let first = parsed
        .spans
        .get(from)
        .ok_or_else(|| bounds(from, parsed.spans.len()))?;
    let second = parsed
        .spans
        .get(to)
        .ok_or_else(|| bounds(to, parsed.spans.len()))?;
    if first.subject != second.subject {
        return invalid("RDF triples can only be reordered within one subject description");
    }
    if from == to {
        return write_graph(package, &path, xml);
    }
    let mut out = String::with_capacity(xml.len());
    if first.start < second.start {
        out.push_str(&xml[..first.start]);
        out.push_str(&xml[first.end..second.end]);
        out.push_str(&xml[first.start..first.end]);
        out.push_str(&xml[second.end..]);
    } else {
        out.push_str(&xml[..second.start]);
        out.push_str(&xml[first.start..first.end]);
        out.push_str(&xml[second.start..first.start]);
        out.push_str(&xml[first.end..]);
    }
    write_graph(package, &path, out)
}

fn parse(path: &str, xml: &str) -> Result<Parsed> {
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
    let mut subject: Option<(usize, OdfRdfSubject)> = None;
    let mut property: Option<(usize, usize, String, OdfRdfObject)> = None;
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
                    for raw in element.attributes().with_checks(true) {
                        let raw = raw.map_err(|error| {
                            Error::InvalidFormat(format!("invalid RDF root attribute: {error}"))
                        })?;
                        let key = std::str::from_utf8(raw.key.as_ref()).unwrap_or("");
                        let value = decode_attr(&reader, &raw)?;
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
                        parse_object_attrs(&reader, &element)?.unwrap_or(OdfRdfObject::Literal {
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
                    && root_depth.is_some()
                    && depth == root_depth.unwrap() + 1
                    && ns.as_deref() == Some(RDF_NS)
                    && element.local_name().as_ref() == b"Description"
                {
                    let _ = parse_subject(&reader, &element)?;
                } else if let Some((subject_depth, current)) = &subject
                    && depth == *subject_depth + 1
                {
                    let predicate = predicate_iri(ns.as_deref(), element.local_name().as_ref())?;
                    let object =
                        parse_object_attrs(&reader, &element)?.unwrap_or(OdfRdfObject::Literal {
                            value: String::new(),
                            datatype: None,
                            language: None,
                        });
                    triples.push(OdfRdfTriple {
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
            Event::Text(value) if property.is_some() => {
                let text = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid RDF literal: {error}"))
                    })?;
                if let Some((_, _, _, OdfRdfObject::Literal { value, .. })) = property.as_mut() {
                    value.push_str(&text);
                    if value.len() > MAX_VALUE {
                        return invalid("RDF literal exceeds size limit");
                    }
                } else if !text.trim().is_empty() {
                    return invalid("RDF resource property cannot contain text");
                }
            },
            Event::CData(value) if property.is_some() => {
                let text = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid RDF literal: {error}"))
                    })?;
                if let Some((_, _, _, OdfRdfObject::Literal { value, .. })) = property.as_mut() {
                    value.push_str(&text);
                    if value.len() > MAX_VALUE {
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
                        property.take().expect("active RDF property");
                    let current = subject.as_ref().expect("RDF subject").1.clone();
                    triples.push(OdfRdfTriple {
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
            _ => {},
        }
        if triples.len() > MAX_TRIPLES {
            return invalid("RDF graph exceeds triple limit");
        }
        buffer.clear();
    }
    let root_close =
        root_close.ok_or_else(|| Error::InvalidFormat("unterminated rdf:RDF root".to_string()))?;
    Ok(Parsed {
        graph: OdfRdfGraph {
            path: path.to_string(),
            base,
            prefixes,
            triples,
        },
        spans,
        root_close,
    })
}

fn parse_subject(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<OdfRdfSubject> {
    let about = rdf_attr(reader, element, b"about")?;
    let node = rdf_attr(reader, element, b"nodeID")?;
    match (about, node) {
        (Some(value), None) => Ok(OdfRdfSubject::Iri(value)),
        (None, Some(value)) => Ok(OdfRdfSubject::BlankNode(value)),
        _ => invalid("rdf:Description requires exactly one of rdf:about or rdf:nodeID"),
    }
}

fn parse_object_attrs(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<Option<OdfRdfObject>> {
    let resource = rdf_attr(reader, element, b"resource")?;
    let node = rdf_attr(reader, element, b"nodeID")?;
    let datatype = rdf_attr(reader, element, b"datatype")?;
    let language = namespaced_attr(reader, element, XML_NS, b"lang")?;
    match (resource, node) {
        (Some(value), None) if datatype.is_none() && language.is_none() => {
            Ok(Some(OdfRdfObject::Iri(value)))
        },
        (None, Some(value)) if datatype.is_none() && language.is_none() => {
            Ok(Some(OdfRdfObject::BlankNode(value)))
        },
        (None, None) => Ok(Some(OdfRdfObject::Literal {
            value: String::new(),
            datatype,
            language,
        })),
        _ => invalid("invalid or ambiguous RDF property object attributes"),
    }
}

fn validate_triples(
    package: &OwnedPackage,
    triples: &[OdfRdfTriple],
    new_path: Option<&str>,
) -> Result<()> {
    if triples.len() > MAX_TRIPLES {
        return invalid("RDF graph exceeds triple limit");
    }
    let anchors = xml_ids(package)?;
    for triple in triples {
        match &triple.subject {
            OdfRdfSubject::Iri(value) => validate_reference(package, value, new_path, &anchors)?,
            OdfRdfSubject::BlankNode(value) => validate_blank(value)?,
        }
        validate_iri(&triple.predicate)?;
        match &triple.object {
            OdfRdfObject::Iri(value) => validate_reference(package, value, new_path, &anchors)?,
            OdfRdfObject::BlankNode(value) => validate_blank(value)?,
            OdfRdfObject::Literal {
                value,
                datatype,
                language,
            } => {
                validate_value(value)?;
                if datatype.is_some() && language.is_some() {
                    return invalid("RDF literal cannot have both datatype and language");
                }
                if let Some(value) = datatype {
                    validate_iri(value)?;
                }
                if let Some(value) = language
                    && (value.is_empty()
                        || value.len() > 128
                        || !value
                            .bytes()
                            .all(|byte| byte.is_ascii_alphanumeric() || byte == b'-'))
                {
                    return invalid("invalid RDF language tag");
                }
            },
        }
    }
    Ok(())
}

fn validate_reference(
    package: &OwnedPackage,
    value: &str,
    new_path: Option<&str>,
    anchors: &HashSet<String>,
) -> Result<()> {
    validate_iri(value)?;
    if let Some(id) = value.strip_prefix('#') {
        if !anchors.contains(id) {
            return invalid(format!("RDF reference '#{id}' has no xml:id anchor"));
        }
    } else if !external_iri(value) && !value.is_empty() {
        let path = value.split('#').next().unwrap_or(value);
        let path = safe_path(path)?;
        if Some(path.as_str()) != new_path && !package.has_file(&path)? {
            return invalid(format!("RDF package reference '{value}' is dangling"));
        }
    }
    Ok(())
}

fn write_graph(package: &OwnedPackage, path: &str, xml: String) -> Result<Vec<u8>> {
    let _ = parse(path, &xml)?;
    let content = String::from_utf8(package.get_file(constants::ODF_CONTENT)?)
        .map_err(|_| Error::InvalidFormat("content.xml is not UTF-8".to_string()))?;
    rebuild_package(
        package,
        &content,
        vec![Addition {
            path: path.to_string(),
            bytes: xml.into_bytes(),
            media_type: constants::ODF_MANIFEST_RDF_TYPE.to_string(),
        }],
        Vec::new(),
        vec![path.to_string()],
        Vec::new(),
    )
}

fn serialize_graph(triples: &[OdfRdfTriple]) -> Result<String> {
    let mut out =
        format!("<?xml version=\"1.0\" encoding=\"UTF-8\"?><rdf:RDF xmlns:rdf=\"{RDF}\">");
    for triple in triples {
        out.push_str(&description_xml(triple)?);
    }
    out.push_str("</rdf:RDF>");
    Ok(out)
}
fn description_xml(triple: &OdfRdfTriple) -> Result<String> {
    let mut out = "<rdf:Description".to_string();
    match &triple.subject {
        OdfRdfSubject::Iri(value) => attr(&mut out, "rdf:about", value)?,
        OdfRdfSubject::BlankNode(value) => attr(&mut out, "rdf:nodeID", value)?,
    }
    out.push('>');
    out.push_str(&property_xml(triple)?);
    out.push_str("</rdf:Description>");
    Ok(out)
}
fn property_xml(triple: &OdfRdfTriple) -> Result<String> {
    let (namespace, local) = split_predicate(&triple.predicate)?;
    let mut out = format!("<p:{local} xmlns:p=\"{}\"", escape(&namespace));
    match &triple.object {
        OdfRdfObject::Iri(value) => {
            attr(&mut out, "rdf:resource", value)?;
            out.push_str("/>");
        },
        OdfRdfObject::BlankNode(value) => {
            attr(&mut out, "rdf:nodeID", value)?;
            out.push_str("/>");
        },
        OdfRdfObject::Literal {
            value,
            datatype,
            language,
        } => {
            if let Some(value) = datatype {
                attr(&mut out, "rdf:datatype", value)?;
            }
            if let Some(value) = language {
                attr(&mut out, "xml:lang", value)?;
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

fn existing_graph(package: &OwnedPackage, path: &str) -> Result<String> {
    let path = safe_path(path)?;
    let archive = package.package()?;
    if !archive.has_file(&path)
        || archive.manifest().get_media_type(&path) != Some(constants::ODF_MANIFEST_RDF_TYPE)
    {
        return invalid(format!("'{path}' is not an RDF metadata part"));
    }
    Ok(path)
}
fn graph_xml(package: &OwnedPackage, path: &str) -> Result<String> {
    String::from_utf8(package.get_file(path)?)
        .map_err(|_| Error::InvalidFormat(format!("RDF metadata part '{path}' is not UTF-8")))
}
fn unused_path(package: &OwnedPackage) -> Result<String> {
    for index in 1..=100_000 {
        let path = format!("Metadata/metadata_{index}.rdf");
        if !package.has_file(&path)? {
            return Ok(path);
        }
    }
    invalid("no collision-free RDF metadata path is available")
}
fn safe_path(value: &str) -> Result<String> {
    let path = media::resolve_package_path(value)?;
    if path.is_empty()
        || path.ends_with('/')
        || path == "mimetype"
        || path.starts_with("META-INF/")
        || matches!(
            path.as_str(),
            "content.xml" | "styles.xml" | "meta.xml" | "settings.xml"
        )
    {
        return invalid("unsafe RDF metadata package path");
    }
    Ok(path)
}
fn triple_refers_to(triple: &OdfRdfTriple, path: &str) -> bool {
    matches!(&triple.object, OdfRdfObject::Iri(value) if value.split('#').next() == Some(path))
}
fn xml_ids(package: &OwnedPackage) -> Result<HashSet<String>> {
    let mut result = HashSet::new();
    for path in [constants::ODF_CONTENT, constants::ODF_STYLES] {
        if !package.has_file(path)? {
            continue;
        }
        let bytes = package.get_file(path)?;
        let xml = std::str::from_utf8(&bytes)
            .map_err(|_| Error::InvalidFormat(format!("{path} is not UTF-8")))?;
        let mut reader = NsReader::from_str(xml);
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("invalid {path}: {error}")))?
            {
                Event::Start(element) | Event::Empty(element) => {
                    for attr in element.attributes().with_checks(true) {
                        let attr = attr.map_err(|error| {
                            Error::InvalidFormat(format!("invalid {path} attribute: {error}"))
                        })?;
                        if attr.key.as_ref() == b"xml:id" {
                            result.insert(decode_attr(&reader, &attr)?);
                        }
                    }
                },
                Event::DocType(_) => {
                    return invalid("DTD is prohibited while validating RDF anchors");
                },
                Event::Eof => break,
                _ => {},
            }
            buffer.clear();
        }
    }
    Ok(result)
}
fn validate_blank(value: &str) -> Result<()> {
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
fn validate_iri(value: &str) -> Result<()> {
    validate_value(value)?;
    if value.chars().any(char::is_whitespace) {
        invalid("RDF IRI contains whitespace")
    } else {
        Ok(())
    }
}
fn validate_value(value: &str) -> Result<()> {
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
fn external_iri(value: &str) -> bool {
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
fn predicate_iri(namespace: Option<&[u8]>, local: &[u8]) -> Result<String> {
    let Some(namespace) = namespace else {
        return invalid("RDF property requires a namespace");
    };
    let namespace = std::str::from_utf8(namespace)
        .map_err(|_| Error::InvalidFormat("invalid RDF predicate namespace".to_string()))?;
    let local = std::str::from_utf8(local)
        .map_err(|_| Error::InvalidFormat("invalid RDF predicate name".to_string()))?;
    Ok(format!("{namespace}{local}"))
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
    for raw in element.attributes().with_checks(true) {
        let raw =
            raw.map_err(|error| Error::InvalidFormat(format!("invalid RDF attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(raw.key);
        if matches!(resolved, ResolveResult::Bound(Namespace(value)) if value == namespace)
            && name.as_ref() == local
        {
            if result.is_some() {
                return invalid("duplicate expanded RDF attribute");
            }
            result = Some(decode_attr(reader, &raw)?);
        }
    }
    Ok(result)
}
fn decode_attr(
    reader: &NsReader<&[u8]>,
    attr: &quick_xml::events::attributes::Attribute<'_>,
) -> Result<String> {
    attr.decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
        .map(|value| value.into_owned())
        .map_err(|error| Error::InvalidFormat(format!("invalid RDF attribute value: {error}")))
}
fn resolved_namespace(value: &ResolveResult<'_>) -> Option<Vec<u8>> {
    match value {
        ResolveResult::Bound(Namespace(value)) => Some(value.to_vec()),
        _ => None,
    }
}
fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| Error::InvalidFormat("RDF XML position overflow".to_string()))
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
fn bounds(index: usize, len: usize) -> Error {
    Error::InvalidFormat(format!(
        "RDF triple index {index} is out of bounds for {len} triples"
    ))
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
