//! ODF package orchestration for RDF metadata parts.

use super::codec::{
    MAX_TRIPLES, bounds, decode_attr, description_xml, external_iri, invalid, parse, property_xml,
    serialize_graph, validate_blank, validate_iri, validate_value,
};
use super::model::{Graph, Object, Subject, Triple};
use crate::constants;
use crate::core::OwnedPackage;
use crate::embedded_chart::{Addition, rebuild_package, splice};
use litchi_core::{Error, Result};
use litchi_odf_common::package::resolve_package_path;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use std::collections::HashSet;

pub(crate) fn graphs(package: &OwnedPackage) -> Result<Vec<Graph>> {
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
    triples: &[Triple],
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
    triples: &[Triple],
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
    triple: &Triple,
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
    triple: &Triple,
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
