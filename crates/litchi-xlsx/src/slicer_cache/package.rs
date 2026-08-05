//! OPC ownership for workbook Slicer Cache parts.
//!
//! The cache definition codec remains in the parent module. This layer owns
//! only the workbook extension, relationship, part, and orphan invariants so
//! callers can use the same cache values for semantic validation and package
//! persistence.

use super::{
    Cache, Result, SLICER_CACHE_CONTENT_TYPE, SLICER_CACHE_RELATIONSHIP_TYPE, invalid, parse,
    validate, write,
};
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const INTEGRATION_URI: &str = "{BBE1A952-AA13-448E-AADC-164F8A28A991}";
const MAX_REWRITE_BYTES: usize = 32 * 1024 * 1024;
const MAX_REFERENCES: usize = 65_536;
const MAX_DEPTH: usize = 256;

/// Load every Slicer Cache relationship listed by the workbook integration
/// extension. Cache binaries and retained XML remain inert.
pub fn load_slicer_caches(package: &OpcPackage) -> Result<Vec<Cache>> {
    reject_root_relationship(package)?;
    let workbook = package.main_document_part()?;
    let (core, rel, ids) = integration_ids(workbook.blob())?;
    let _ = (core, rel);
    let mut seen_ids = HashSet::new();
    let mut seen_targets = HashSet::new();
    let mut output = Vec::with_capacity(ids.len());

    for id in ids {
        if !seen_ids.insert(id.clone()) {
            return Err(invalid(format!("duplicate Slicer Cache reference '{id}'")));
        }
        let relationship = workbook
            .rels()
            .get(&id)
            .ok_or_else(|| invalid(format!("missing Slicer Cache relationship '{id}'")))?;
        if relationship.reltype() != SLICER_CACHE_RELATIONSHIP_TYPE || relationship.is_external() {
            return Err(invalid(format!(
                "Slicer Cache reference '{id}' must target an internal Slicer Cache relationship"
            )));
        }
        let target = relationship.target_partname()?;
        if !seen_targets.insert(target.to_string()) {
            return Err(invalid(format!(
                "multiple Slicer Cache references target '{target}'"
            )));
        }
        let part = package.get_part(&target)?;
        if part.content_type() != SLICER_CACHE_CONTENT_TYPE {
            return Err(invalid(format!(
                "Slicer Cache part '{target}' has content type '{}'",
                part.content_type()
            )));
        }
        if !part.rels().is_empty() {
            return Err(invalid(format!(
                "Slicer Cache part '{target}' has forbidden outbound relationships"
            )));
        }
        output.push(Cache {
            relationship_id: id,
            part_name: target.to_string(),
            definition: parse(part.blob())?,
        });
    }

    for relationship in workbook
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == SLICER_CACHE_RELATIONSHIP_TYPE)
    {
        if !seen_ids.contains(relationship.r_id()) {
            return Err(invalid(format!(
                "unreferenced Slicer Cache relationship '{}'",
                relationship.r_id()
            )));
        }
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == SLICER_CACHE_CONTENT_TYPE)
    {
        if !seen_targets.contains(part.partname().as_str()) {
            return Err(invalid(format!(
                "orphan Slicer Cache part '{}'",
                part.partname()
            )));
        }
    }
    validate_set(&output)?;
    Ok(output)
}

/// Add one Slicer Cache part and append its workbook integration reference.
/// The operation is staged on a package clone so malformed graphs cannot
/// partially mutate the caller's package.
pub fn store_slicer_cache(package: &mut OpcPackage, value: &Cache) -> Result<()> {
    validate(&value.definition)?;
    let mut staged = package.clone();
    store_slicer_cache_in_place(&mut staged, value)?;
    *package = staged;
    Ok(())
}

fn store_slicer_cache_in_place(package: &mut OpcPackage, value: &Cache) -> Result<()> {
    let workbook_name = package.main_document_part()?.partname().clone();
    let existing = load_slicer_caches(package)?;
    if existing.iter().any(|cache| {
        cache
            .definition
            .name
            .eq_ignore_ascii_case(&value.definition.name)
    }) {
        return Err(invalid(format!(
            "duplicate Slicer Cache name '{}'",
            value.definition.name
        )));
    }
    validate_set(
        &existing
            .iter()
            .cloned()
            .chain(std::iter::once(value.clone()))
            .collect::<Vec<_>>(),
    )?;

    validate_relationship_id(&value.relationship_id)?;
    let uri = PackURI::new(&value.part_name)
        .map_err(|error| invalid(format!("invalid Slicer Cache part URI: {error}")))?;
    if !uri.as_str().starts_with("/xl/slicerCaches/") || !uri.as_str().ends_with(".xml") {
        return Err(invalid(format!(
            "Slicer Cache part '{uri}' must be under /xl/slicerCaches and end in .xml"
        )));
    }
    let workbook = package.get_part(&workbook_name)?;
    if workbook.rels().get(&value.relationship_id).is_some() {
        return Err(invalid(format!(
            "workbook relationship ID '{}' already exists",
            value.relationship_id
        )));
    }
    if package.get_part(&uri).is_ok() {
        return Err(invalid(format!(
            "Slicer Cache target '{uri}' already exists"
        )));
    }
    let xml = write(&value.definition)?;
    let ids = existing
        .iter()
        .map(|cache| cache.relationship_id.clone())
        .chain(std::iter::once(value.relationship_id.clone()))
        .collect::<Vec<_>>();
    let updated = rewrite_or_insert(
        workbook.blob(),
        &ids,
        INTEGRATION_URI,
        "slicerCaches",
        "slicerCache",
        X14,
    )?;
    package.try_add_part(Box::new(BlobPart::new(
        uri.clone(),
        SLICER_CACHE_CONTENT_TYPE.into(),
        xml,
    )))?;
    package
        .get_part_mut(&workbook_name)?
        .rels_mut()
        .add_relationship(
            SLICER_CACHE_RELATIONSHIP_TYPE.into(),
            uri.relative_ref(workbook_name.base_uri()),
            value.relationship_id.clone(),
            false,
        );
    package.get_part_mut(&workbook_name)?.set_blob(updated);
    package.unsign();
    Ok(())
}

fn validate_set(caches: &[Cache]) -> Result<()> {
    if caches.len() > MAX_REFERENCES {
        return Err(invalid("Slicer Cache reference count exceeds the limit"));
    }
    let mut names = HashSet::new();
    let mut uids = HashSet::new();
    let all_uids = caches.iter().any(|cache| cache.definition.uid.is_some());
    for cache in caches {
        validate(&cache.definition)?;
        validate_relationship_id(&cache.relationship_id)?;
        if !names.insert(cache.definition.name.to_ascii_lowercase()) {
            return Err(invalid("duplicate Slicer Cache name"));
        }
        match &cache.definition.uid {
            Some(uid) if !uids.insert(uid.to_ascii_lowercase()) => {
                return Err(invalid("duplicate Slicer Cache uid"));
            },
            None if all_uids => {
                return Err(invalid(
                    "Slicer Cache uid must be present on every cache or none",
                ));
            },
            _ => {},
        }
    }
    Ok(())
}

fn reject_root_relationship(package: &OpcPackage) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == SLICER_CACHE_RELATIONSHIP_TYPE)
    {
        return Err(invalid(
            "package root cannot source a Slicer Cache relationship",
        ));
    }
    for part in package.iter_parts() {
        if part.partname() == package.main_document_part()?.partname() {
            continue;
        }
        if part
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == SLICER_CACHE_RELATIONSHIP_TYPE)
        {
            return Err(invalid(format!(
                "non-workbook part '{}' sources a Slicer Cache relationship",
                part.partname()
            )));
        }
    }
    Ok(())
}

fn integration_ids(xml: &[u8]) -> Result<(&'static str, &'static str, Vec<String>)> {
    let (core, rel) = root_namespaces(xml)?;
    let mut reader = NsReader::from_reader(xml);
    let decoder = reader.decoder();
    let mut depth = 0usize;
    let mut in_extension = false;
    let mut in_caches = false;
    let mut extension_seen = false;
    let mut ids = Vec::new();
    loop {
        let event = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("Slicer Cache integration XML: {error}")))?;
        let (namespace, event) = event;
        match event {
            Event::Start(element) => {
                if depth == 2
                    && exact(&namespace, core)
                    && element.local_name().as_ref() == b"ext"
                    && attribute(&element, b"uri", decoder)?.as_deref() == Some(INTEGRATION_URI)
                {
                    if extension_seen {
                        return Err(invalid("duplicate Slicer Cache integration extension"));
                    }
                    extension_seen = true;
                    in_extension = true;
                } else if in_extension
                    && exact(&namespace, X14)
                    && element.local_name().as_ref() == b"slicerCaches"
                {
                    if in_caches {
                        return Err(invalid("duplicate slicerCaches element"));
                    }
                    in_caches = true;
                } else if in_caches
                    && exact(&namespace, X14)
                    && element.local_name().as_ref() == b"slicerCache"
                {
                    return Err(invalid("non-empty slicerCache element is not supported"));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XML depth overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(invalid("Slicer Cache integration XML is too deep"));
                }
            },
            Event::Empty(element) => {
                if in_caches
                    && exact(&namespace, X14)
                    && element.local_name().as_ref() == b"slicerCache"
                {
                    let id = attribute(&element, b"r:id", reader.decoder())?
                        .ok_or_else(|| invalid("slicerCache is missing r:id"))?;
                    ids.push(id);
                    if ids.len() > MAX_REFERENCES {
                        return Err(invalid("Slicer Cache reference count exceeds the limit"));
                    }
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid(
                        "unexpected Slicer Cache integration closing element",
                    ));
                }
                depth -= 1;
                if in_caches
                    && exact(&namespace, X14)
                    && element.local_name().as_ref() == b"slicerCaches"
                {
                    in_caches = false;
                } else if in_extension
                    && exact(&namespace, core)
                    && element.local_name().as_ref() == b"ext"
                {
                    in_extension = false;
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || in_extension || in_caches {
        return Err(invalid("incomplete Slicer Cache integration XML"));
    }
    Ok((core, rel, ids))
}

fn rewrite_or_insert(
    xml: &[u8],
    ids: &[String],
    uri: &str,
    list: &str,
    item: &str,
    family_ns: &str,
) -> Result<Vec<u8>> {
    let (core, rel) = root_namespaces(xml)?;
    let mut reader = NsReader::from_reader(xml);
    let decoder = reader.decoder();
    let mut depth = 0usize;
    let mut start = None;
    let mut end = None;
    loop {
        let offset = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("Slicer Cache XML offset overflow"))?;
        let event = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("Slicer Cache integration XML: {error}")))?;
        let (namespace, event) = event;
        match event {
            Event::Start(element)
                if depth == 2
                    && exact(&namespace, core)
                    && element.local_name().as_ref() == b"ext"
                    && attribute(&element, b"uri", decoder)?.as_deref() == Some(uri) =>
            {
                if start.is_some() {
                    return Err(invalid("duplicate integration extension"));
                }
                start = Some(offset);
                depth += 1;
            },
            Event::Empty(element)
                if depth == 2
                    && exact(&namespace, core)
                    && element.local_name().as_ref() == b"ext"
                    && attribute(&element, b"uri", decoder)?.as_deref() == Some(uri) =>
            {
                if start.is_some() {
                    return Err(invalid("duplicate integration extension"));
                }
                start = Some(offset);
                end = Some(
                    usize::try_from(reader.buffer_position())
                        .map_err(|_| invalid("Slicer Cache XML offset overflow"))?,
                );
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected Slicer Cache XML close"));
                }
                depth -= 1;
                if start.is_some()
                    && depth == 2
                    && exact(&namespace, core)
                    && element.local_name().as_ref() == b"ext"
                {
                    end = Some(
                        usize::try_from(reader.buffer_position())
                            .map_err(|_| invalid("Slicer Cache XML offset overflow"))?,
                    );
                }
            },
            Event::Start(_) => depth += 1,
            Event::Empty(_) => {},
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
        if depth > MAX_DEPTH {
            return Err(invalid("Slicer Cache integration XML is too deep"));
        }
    }
    let fragment = integration_fragment(core, rel, uri, family_ns, list, item, ids);
    if let (Some(start), Some(end)) = (start, end) {
        let size = xml
            .len()
            .checked_sub(end.saturating_sub(start))
            .and_then(|value| value.checked_add(fragment.len()))
            .ok_or_else(|| invalid("Slicer Cache rewrite size overflow"))?;
        if size > MAX_REWRITE_BYTES {
            return Err(invalid("rewritten Slicer Cache XML exceeds limit"));
        }
        let mut output = Vec::with_capacity(size);
        output.extend_from_slice(&xml[..start]);
        output.extend_from_slice(&fragment);
        output.extend_from_slice(&xml[end..]);
        return Ok(output);
    }
    let close = xml
        .windows(b"</workbook>".len())
        .rposition(|window| window == b"</workbook>")
        .ok_or_else(|| invalid("workbook closing element is missing"))?;
    let size = xml
        .len()
        .checked_add(fragment.len())
        .ok_or_else(|| invalid("Slicer Cache XML size overflow"))?;
    if size > MAX_REWRITE_BYTES {
        return Err(invalid("rewritten Slicer Cache XML exceeds limit"));
    }
    let mut output = Vec::with_capacity(size);
    output.extend_from_slice(&xml[..close]);
    output.extend_from_slice(b"<extLst>");
    output.extend_from_slice(&fragment);
    output.extend_from_slice(b"</extLst>");
    output.extend_from_slice(&xml[close..]);
    Ok(output)
}

fn integration_fragment(
    core: &str,
    rel: &str,
    uri: &str,
    family_ns: &str,
    list: &str,
    item: &str,
    ids: &[String],
) -> Vec<u8> {
    let mut output = format!(
        "<ext xmlns=\"{core}\" uri=\"{uri}\"><f:{list} xmlns:f=\"{family_ns}\" xmlns:r=\"{rel}\">"
    );
    for id in ids {
        output.push_str(&format!("<f:{item} r:id=\"{}\"/>", xml_escape(id)));
    }
    output.push_str(&format!("</f:{list}></ext>"));
    output.into_bytes()
}

fn root_namespaces(xml: &[u8]) -> Result<(&'static str, &'static str)> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("workbook XML: {error}")))?;
        match event {
            Event::Start(element) => {
                if element.local_name().as_ref() != b"workbook" {
                    return Err(invalid("workbook XML has an invalid root"));
                }
                return match namespace {
                    ResolveResult::Bound(Namespace(value)) if value == SML.as_bytes() => {
                        Ok((SML, REL))
                    },
                    ResolveResult::Bound(Namespace(value)) if value == STRICT_SML.as_bytes() => {
                        Ok((STRICT_SML, STRICT_REL))
                    },
                    _ => Err(invalid("unsupported SpreadsheetML workbook namespace")),
                };
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => return Err(invalid("workbook XML root is missing")),
            _ => {},
        }
    }
}

fn attribute(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<String>> {
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(|error| invalid(error.to_string()))?;
        if item.key.as_ref() == name {
            return Ok(Some(
                item.decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(|error| invalid(error.to_string()))?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn exact(namespace: &ResolveResult<'_>, value: &str) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(namespace)) if *namespace == value.as_bytes())
}

fn validate_relationship_id(value: &str) -> Result<()> {
    if value.is_empty()
        || !value.bytes().enumerate().all(|(index, byte)| {
            (index == 0 && (byte.is_ascii_alphabetic() || byte == b'_'))
                || (index > 0
                    && (byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.')))
        })
    {
        return Err(invalid(format!(
            "invalid Slicer Cache relationship ID '{value}'"
        )));
    }
    Ok(())
}

fn xml_escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
}
