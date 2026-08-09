//! ODF package orchestration for RDF metadata parts.

use super::codec::{
    MAX_TRIPLES, bounds, decode_attr, description_xml, external_iri, invalid, parse, property_xml,
    serialize_graph, validate_blank, validate_iri, validate_value,
};
use super::model::{Graph, Object, Subject, Triple};
use crate::constants;
use crate::core::OwnedPackage;
use crate::package::resolve_package_path;
use crate::package::{Addition, rebuild_package, splice};
use litchi_core::{Error, Result};
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use std::collections::HashSet;

/// Read every RDF metadata graph declared by an ODF package manifest.
///
/// # Errors
///
/// Returns an error when the package, manifest, or an RDF metadata part is malformed.
pub fn graphs(package: &OwnedPackage) -> Result<Vec<Graph>> {
    let archive = package.package()?;
    let mut paths: Vec<String> = archive
        .manifest()
        .entries
        .iter()
        .filter(|(path, entry)| {
            entry.media_type == constants::ODF_MANIFEST_RDF_TYPE && !path.ends_with('/')
        })
        .map(|(path, _)| path.clone())
        .collect();
    paths.sort();
    let mut result = Vec::with_capacity(paths.len());
    for path in paths {
        if !archive.has_file(&path) {
            return invalid(format!("RDF manifest entry '{path}' is dangling"));
        }
        let xml = String::from_utf8(archive.get_file(&path)?).map_err(|error| {
            Error::InvalidFormat(format!("RDF metadata part '{path}' is not UTF-8: {error}"))
        })?;
        result.push(parse(&path, &xml)?.graph);
    }
    Ok(result)
}

/// Add a validated RDF metadata graph to an ODF package.
///
/// # Errors
///
/// Returns an error when the requested path, triples, or package is invalid.
pub fn add_graph(
    package: &OwnedPackage,
    preferred: Option<&str>,
    triples: &[Triple],
) -> Result<(Vec<u8>, String)> {
    let path = match preferred {
        Some(preferred_path) => {
            let graph_path = safe_path(preferred_path)?;
            if package.has_file(&graph_path)? {
                return invalid(format!("RDF metadata path '{graph_path}' already exists"));
            }
            graph_path
        },
        None => unused_path(package)?,
    };
    validate_triples(package, triples, Some(&path))?;
    let xml = serialize_graph(triples)?;
    let content = String::from_utf8(package.get_file(constants::ODF_CONTENT)?)
        .map_err(|error| Error::InvalidFormat(format!("content.xml is not UTF-8: {error}")))?;
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

/// Replace all triples in one RDF metadata graph.
///
/// # Errors
///
/// Returns an error when the graph path, triples, or package is invalid.
pub fn replace_graph(package: &OwnedPackage, path: &str, triples: &[Triple]) -> Result<Vec<u8>> {
    let graph_path = existing_graph(package, path)?;
    validate_triples(package, triples, Some(&graph_path))?;
    write_graph(package, &graph_path, serialize_graph(triples)?)
}

/// Remove an RDF metadata graph when no other graph references it.
///
/// # Errors
///
/// Returns an error when the graph does not exist, remains referenced, or the package is invalid.
pub fn remove_graph(package: &OwnedPackage, path: &str) -> Result<Vec<u8>> {
    let graph_path = existing_graph(package, path)?;
    for graph in graphs(package)? {
        if graph.path != graph_path
            && graph
                .triples
                .iter()
                .any(|triple| triple_refers_to(triple, &graph_path))
        {
            return invalid(format!(
                "RDF metadata part '{graph_path}' is still referenced by '{}'",
                graph.path
            ));
        }
    }
    let content = String::from_utf8(package.get_file(constants::ODF_CONTENT)?)
        .map_err(|error| Error::InvalidFormat(format!("content.xml is not UTF-8: {error}")))?;
    rebuild_package(
        package,
        &content,
        Vec::new(),
        Vec::new(),
        vec![graph_path],
        Vec::new(),
    )
}

/// Append one triple to an RDF metadata graph.
///
/// # Errors
///
/// Returns an error when the graph, triple, or package is invalid.
pub fn add_triple(package: &OwnedPackage, path: &str, triple: &Triple) -> Result<(Vec<u8>, usize)> {
    let graph_path = existing_graph(package, path)?;
    validate_triples(package, std::slice::from_ref(triple), Some(&graph_path))?;
    let xml = graph_xml(package, &graph_path)?;
    let parsed = parse(&graph_path, &xml)?;
    if parsed.graph.triples.len() >= MAX_TRIPLES {
        return invalid("RDF graph exceeds triple limit");
    }
    let fragment = description_xml(triple)?;
    let updated = splice(&xml, parsed.root_close, parsed.root_close, &fragment)?;
    write_graph(package, &graph_path, updated).map(|bytes| (bytes, parsed.graph.triples.len()))
}

/// Replace one triple while retaining its RDF description subject.
///
/// # Errors
///
/// Returns an error when the index, triple, graph, or package is invalid.
pub fn replace_triple(
    package: &OwnedPackage,
    path: &str,
    index: usize,
    triple: &Triple,
) -> Result<Vec<u8>> {
    let graph_path = existing_graph(package, path)?;
    validate_triples(package, std::slice::from_ref(triple), Some(&graph_path))?;
    let xml = graph_xml(package, &graph_path)?;
    let parsed = parse(&graph_path, &xml)?;
    let span = parsed
        .spans
        .get(index)
        .ok_or_else(|| bounds(index, parsed.spans.len()))?;
    if triple.subject != span.subject {
        return invalid("replacing an RDF property cannot change its description subject");
    }
    let updated = splice(&xml, span.start, span.end, &property_xml(triple)?)?;
    write_graph(package, &graph_path, updated)
}

/// Remove one triple from an RDF metadata graph.
///
/// # Errors
///
/// Returns an error when the index, graph, or package is invalid.
pub fn remove_triple(package: &OwnedPackage, path: &str, index: usize) -> Result<Vec<u8>> {
    let graph_path = existing_graph(package, path)?;
    let xml = graph_xml(package, &graph_path)?;
    let parsed = parse(&graph_path, &xml)?;
    let span = parsed
        .spans
        .get(index)
        .ok_or_else(|| bounds(index, parsed.spans.len()))?;
    write_graph(
        package,
        &graph_path,
        splice(&xml, span.start, span.end, "")?,
    )
}

/// Reorder two triples that share an RDF description subject.
///
/// # Errors
///
/// Returns an error when either index is invalid, the triples have different subjects, or the
/// package cannot be rewritten.
pub fn move_triple(package: &OwnedPackage, path: &str, from: usize, to: usize) -> Result<Vec<u8>> {
    let graph_path = existing_graph(package, path)?;
    let xml = graph_xml(package, &graph_path)?;
    let parsed = parse(&graph_path, &xml)?;
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
        return write_graph(package, &graph_path, xml);
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
    write_graph(package, &graph_path, out)
}

fn validate_triples(
    package: &OwnedPackage,
    triples: &[Triple],
    new_path: Option<&str>,
) -> Result<()> {
    if triples.len() > MAX_TRIPLES {
        return invalid("RDF graph exceeds triple limit");
    }
    let anchors = xml_ids(package)?;
    for triple in triples {
        match &triple.subject {
            Subject::Iri(value) => validate_reference(package, value, new_path, &anchors)?,
            Subject::BlankNode(value) => validate_blank(value)?,
        }
        validate_iri(&triple.predicate)?;
        match &triple.object {
            Object::Iri(value) => validate_reference(package, value, new_path, &anchors)?,
            Object::BlankNode(value) => validate_blank(value)?,
            Object::Literal {
                value: literal_value,
                datatype,
                language,
            } => {
                validate_value(literal_value)?;
                if datatype.is_some() && language.is_some() {
                    return invalid("RDF literal cannot have both datatype and language");
                }
                if let Some(datatype_iri) = datatype {
                    validate_iri(datatype_iri)?;
                }
                if let Some(language_tag) = language
                    && (language_tag.is_empty()
                        || language_tag.len() > 128
                        || !language_tag
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
        let reference_path = value.split('#').next().unwrap_or(value);
        let resolved_path = safe_path(reference_path)?;
        if Some(resolved_path.as_str()) != new_path && !package.has_file(&resolved_path)? {
            return invalid(format!("RDF package reference '{value}' is dangling"));
        }
    }
    Ok(())
}

fn write_graph(package: &OwnedPackage, path: &str, xml: String) -> Result<Vec<u8>> {
    let _ = parse(path, &xml)?;
    let content = String::from_utf8(package.get_file(constants::ODF_CONTENT)?)
        .map_err(|error| Error::InvalidFormat(format!("content.xml is not UTF-8: {error}")))?;
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

fn existing_graph(package: &OwnedPackage, path: &str) -> Result<String> {
    let graph_path = safe_path(path)?;
    let archive = package.package()?;
    if !archive.has_file(&graph_path)
        || archive.manifest().get_media_type(&graph_path) != Some(constants::ODF_MANIFEST_RDF_TYPE)
    {
        return invalid(format!("'{graph_path}' is not an RDF metadata part"));
    }
    Ok(graph_path)
}
fn graph_xml(package: &OwnedPackage, path: &str) -> Result<String> {
    String::from_utf8(package.get_file(path)?).map_err(|error| {
        Error::InvalidFormat(format!("RDF metadata part '{path}' is not UTF-8: {error}"))
    })
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
    let path = resolve_package_path(value)?;
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
fn triple_refers_to(triple: &Triple, path: &str) -> bool {
    matches!(&triple.object, Object::Iri(value) if value.split('#').next() == Some(path))
}
fn xml_ids(package: &OwnedPackage) -> Result<HashSet<String>> {
    let mut result = HashSet::new();
    for path in [constants::ODF_CONTENT, constants::ODF_STYLES] {
        if !package.has_file(path)? {
            continue;
        }
        let bytes = package.get_file(path)?;
        let xml = std::str::from_utf8(&bytes)
            .map_err(|error| Error::InvalidFormat(format!("{path} is not UTF-8: {error}")))?;
        let mut reader = NsReader::from_str(xml);
        let mut buffer = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("invalid {path}: {error}")))?
            {
                Event::Start(element) | Event::Empty(element) => {
                    for raw_attribute in element.attributes().with_checks(true) {
                        let attribute = raw_attribute.map_err(|error| {
                            Error::InvalidFormat(format!("invalid {path} attribute: {error}"))
                        })?;
                        if attribute.key.as_ref() == b"xml:id" {
                            result.insert(decode_attr(&reader, &attribute)?);
                        }
                    }
                },
                Event::DocType(_) => {
                    return invalid("DTD is prohibited while validating RDF anchors");
                },
                Event::Eof => break,
                Event::End(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::GeneralRef(_) => {},
            }
            buffer.clear();
        }
    }
    Ok(result)
}
