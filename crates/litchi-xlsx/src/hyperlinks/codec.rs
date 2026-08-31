use std::collections::{HashMap, HashSet};

use litchi_opc::Relationships;
use litchi_opc::constants::relationship_type as rt;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::{NsReader, Reader};

use super::model::{Hyperlink, HyperlinkReference, validate_text};
use crate::error::{Result, allocation, invalid};

const TRANSITIONAL_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_MAIN: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const TRANSITIONAL_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MAX_WORKSHEET_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_HYPERLINKS: usize = 1_048_576;
const MAX_XML_DEPTH: usize = 64;
const MAX_XML_EVENTS: usize = 4_194_304;
const MAX_XML_NAME_BYTES: usize = 1_024;
const MAX_RELATIONSHIP_ID_BYTES: usize = 1_024;

#[derive(Clone, Copy)]
enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    const fn main_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_MAIN,
            Self::Strict => STRICT_MAIN,
        }
    }

    const fn relationship_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_REL,
            Self::Strict => STRICT_REL,
        }
    }

    const fn hyperlink_relationship_type(self) -> &'static str {
        match self {
            Self::Transitional => rt::HYPERLINK,
            Self::Strict => rt::STRICT_HYPERLINK,
        }
    }
}

pub(crate) fn parse(xml: &[u8], relationships: &Relationships) -> Result<Vec<Hyperlink>> {
    parse_with_event_limit(xml, relationships, MAX_XML_EVENTS)
}

/// One parsed hyperlink together with its private relationship identity.
/// Relationship IDs are needed only while staging a package mutation and do
/// not cross the public semantic boundary.
#[derive(Debug, Clone)]
pub(crate) struct ParsedHyperlink {
    pub(crate) value: Hyperlink,
    pub(crate) relationship_id: Option<String>,
}

pub(crate) fn parse_with_relationship_ids(
    xml: &[u8],
    relationships: &Relationships,
) -> Result<Vec<ParsedHyperlink>> {
    let values = parse(xml, relationships)?;
    let ids = relationship_ids(xml)?;
    if values.len() != ids.len() {
        return Err(invalid(
            "XLSX hyperlink relationship projection is inconsistent",
        ));
    }
    validate_exclusive_relationship_references(xml, &ids)?;
    Ok(values
        .into_iter()
        .zip(ids)
        .map(|(value, relationship_id)| ParsedHyperlink {
            value,
            relationship_id,
        })
        .collect())
}

fn parse_with_event_limit(
    xml: &[u8],
    relationships: &Relationships,
    max_events: usize,
) -> Result<Vec<Hyperlink>> {
    if xml.len() > MAX_WORKSHEET_XML_BYTES {
        return Err(invalid(format!(
            "XLSX worksheet XML exceeds the {MAX_WORKSHEET_XML_BYTES} byte safety limit"
        )));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = false;
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut event_count = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut conformance = None;
    let mut seen_container = false;
    let mut container_depth = None;
    let mut open_hyperlink_depth = None;
    let mut hyperlinks = Vec::new();
    let mut ranges = HashSet::new();
    let mut used_relationships = HashSet::<String>::new();
    let mut open_names = Vec::<Box<[u8]>>::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid XLSX worksheet hyperlink XML: {error}")))?;
        if !matches!(event, Event::Eof) {
            event_count = event_count
                .checked_add(1)
                .ok_or_else(|| invalid("XLSX worksheet XML event count overflows usize"))?;
            if event_count > max_events {
                return Err(invalid(format!(
                    "XLSX worksheet hyperlink XML exceeds the {max_events} event limit"
                )));
            }
        }
        match event {
            Event::Start(element) => {
                let local = decode_element_local(&element)?;
                let namespace = resolve_namespace(&namespace)?;
                open_names.push(element.name().as_ref().to_vec().into_boxed_slice());
                if depth == 0 {
                    if root_seen || root_closed {
                        return Err(invalid("XLSX worksheet XML has more than one root"));
                    }
                    let parsed = parse_root(namespace.as_deref(), &local)?;
                    conformance = Some(parsed);
                    root_seen = true;
                } else if depth == 1 && local == "hyperlinks" {
                    let parsed =
                        conformance.ok_or_else(|| invalid("XLSX worksheet root is missing"))?;
                    require_name(namespace.as_deref(), &local, parsed, "hyperlinks")?;
                    if seen_container {
                        return Err(invalid(
                            "XLSX worksheet has duplicate hyperlinks containers",
                        ));
                    }
                    validate_container_attributes(&reader, &element)?;
                    seen_container = true;
                    container_depth = Some(depth + 1);
                } else if container_depth == Some(depth) {
                    let parsed =
                        conformance.ok_or_else(|| invalid("XLSX worksheet root is missing"))?;
                    require_name(namespace.as_deref(), &local, parsed, "hyperlink")?;
                    push_hyperlink(
                        &reader,
                        &element,
                        parsed,
                        relationships,
                        &mut used_relationships,
                        &mut ranges,
                        &mut hyperlinks,
                    )?;
                    open_hyperlink_depth = Some(depth + 1);
                } else if container_depth.is_some() && depth > container_depth.unwrap_or(0) {
                    return Err(invalid(
                        "XLSX hyperlink owner contains unsupported child markup",
                    ));
                } else if matches!(local.as_str(), "hyperlinks" | "hyperlink")
                    && namespace.as_deref().is_some_and(is_main_namespace)
                {
                    return Err(invalid(
                        "XLSX hyperlink owner is not a direct worksheet child",
                    ));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XLSX worksheet XML depth overflows usize"))?;
                if depth > MAX_XML_DEPTH {
                    return Err(invalid(format!(
                        "XLSX worksheet hyperlink XML exceeds the {MAX_XML_DEPTH} depth limit"
                    )));
                }
            },
            Event::Empty(element) => {
                let local = decode_element_local(&element)?;
                let namespace = resolve_namespace(&namespace)?;
                if depth == 0 {
                    if root_seen || root_closed {
                        return Err(invalid("XLSX worksheet XML has more than one root"));
                    }
                    conformance = Some(parse_root(namespace.as_deref(), &local)?);
                    root_seen = true;
                    root_closed = true;
                } else if depth == 1 && local == "hyperlinks" {
                    let parsed =
                        conformance.ok_or_else(|| invalid("XLSX worksheet root is missing"))?;
                    require_name(namespace.as_deref(), &local, parsed, "hyperlinks")?;
                    if seen_container {
                        return Err(invalid(
                            "XLSX worksheet has duplicate hyperlinks containers",
                        ));
                    }
                    validate_container_attributes(&reader, &element)?;
                    seen_container = true;
                } else if container_depth == Some(depth) {
                    let parsed =
                        conformance.ok_or_else(|| invalid("XLSX worksheet root is missing"))?;
                    require_name(namespace.as_deref(), &local, parsed, "hyperlink")?;
                    push_hyperlink(
                        &reader,
                        &element,
                        parsed,
                        relationships,
                        &mut used_relationships,
                        &mut ranges,
                        &mut hyperlinks,
                    )?;
                } else if container_depth.is_some() || open_hyperlink_depth.is_some() {
                    return Err(invalid(
                        "XLSX hyperlink owner contains unsupported child markup",
                    ));
                } else if matches!(local.as_str(), "hyperlinks" | "hyperlink")
                    && namespace.as_deref().is_some_and(is_main_namespace)
                {
                    return Err(invalid(
                        "XLSX hyperlink owner is not a direct worksheet child",
                    ));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("XLSX worksheet XML depth underflows usize"));
                }
                validate_raw_name(element.name().as_ref(), "end element name")?;
                let expected = open_names
                    .pop()
                    .ok_or_else(|| invalid("XLSX worksheet end element has no matching start"))?;
                if expected.as_ref() != element.name().as_ref() {
                    return Err(invalid("XLSX worksheet start and end element names differ"));
                }
                if open_hyperlink_depth == Some(depth) {
                    open_hyperlink_depth = None;
                }
                if container_depth == Some(depth) {
                    container_depth = None;
                }
                if depth == 1 {
                    root_closed = true;
                }
                depth -= 1;
            },
            Event::Text(text) if depth == 0 => {
                if !text
                    .as_ref()
                    .iter()
                    .all(|byte| matches!(byte, b' ' | b'\t' | b'\n' | b'\r'))
                {
                    return Err(invalid("XLSX worksheet XML has text outside its root"));
                }
            },
            Event::Text(text) if depth == 1 => {
                if !text
                    .as_ref()
                    .iter()
                    .all(|byte| matches!(byte, b' ' | b'\t' | b'\n' | b'\r'))
                {
                    return Err(invalid(
                        "XLSX worksheet XML has text outside a worksheet child",
                    ));
                }
            },
            Event::Text(text) if container_depth.is_some() => {
                if !text
                    .as_ref()
                    .iter()
                    .all(|byte| matches!(byte, b' ' | b'\t' | b'\n' | b'\r'))
                {
                    return Err(invalid(
                        "XLSX hyperlink owner contains unsupported text content",
                    ));
                }
            },
            Event::Comment(_) | Event::PI(_) | Event::CData(_) | Event::GeneralRef(_)
                if container_depth.is_some() =>
            {
                return Err(invalid(
                    "XLSX hyperlink owner contains unsupported opaque markup",
                ));
            },
            Event::DocType(_) => {
                return Err(invalid("XLSX worksheet hyperlink XML cannot contain a DTD"));
            },
            Event::CData(_) | Event::GeneralRef(_) if depth <= 1 => {
                return Err(invalid(
                    "XLSX worksheet XML has character data outside its root",
                ));
            },
            Event::Decl(_) if root_seen => {
                return Err(invalid(
                    "XLSX worksheet XML declaration must precede the root",
                ));
            },
            Event::Eof => break,
            Event::Decl(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_)
            | Event::Comment(_)
            | Event::PI(_) => {},
        }
    }

    let conformance = conformance.ok_or_else(|| invalid("XLSX worksheet root is missing"))?;
    if !root_seen
        || !root_closed
        || depth != 0
        || !open_names.is_empty()
        || container_depth.is_some()
        || open_hyperlink_depth.is_some()
    {
        return Err(invalid("XLSX worksheet hyperlink XML is not balanced"));
    }
    for relationship in relationships.iter().filter(|relationship| {
        matches!(relationship.reltype(), rt::HYPERLINK | rt::STRICT_HYPERLINK)
    }) {
        if relationship.reltype() != conformance.hyperlink_relationship_type() {
            return Err(invalid(
                "XLSX hyperlink relationship does not match the worksheet conformance",
            ));
        }
        if !used_relationships.contains(relationship.r_id()) {
            return Err(invalid(
                "XLSX worksheet has an orphan hyperlink relationship",
            ));
        }
    }
    Ok(hyperlinks)
}

fn parse_root(namespace: Option<&str>, local: &str) -> Result<Conformance> {
    if local != "worksheet" {
        return Err(invalid(format!(
            "XLSX hyperlink projection requires a worksheet root, found '{local}'"
        )));
    }
    match namespace {
        Some(TRANSITIONAL_MAIN) => Ok(Conformance::Transitional),
        Some(STRICT_MAIN) => Ok(Conformance::Strict),
        _ => Err(invalid("XLSX worksheet uses an unsupported namespace")),
    }
}

fn require_name(
    namespace: Option<&str>,
    local: &str,
    conformance: Conformance,
    expected_local: &str,
) -> Result<()> {
    if namespace == Some(conformance.main_namespace()) && local == expected_local {
        Ok(())
    } else {
        Err(invalid(format!(
            "XLSX hyperlink owner contains unexpected element '{local}'"
        )))
    }
}

fn validate_container_attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid(format!("invalid XLSX hyperlinks attribute: {error}")))?;
        validate_raw_name(attribute.key.as_ref(), "hyperlinks attribute name")?;
        let name = attribute.key.as_ref();
        if name == b"xmlns" || name.starts_with(b"xmlns:") {
            continue;
        }
        let (_, local) = reader.resolver().resolve_attribute(attribute.key);
        let local = decode(local.as_ref(), "hyperlinks attribute local name")?;
        return Err(invalid(format!(
            "XLSX hyperlinks container has unsupported attribute '{local}'"
        )));
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_hyperlink(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    conformance: Conformance,
    relationships: &Relationships,
    used_relationships: &mut HashSet<String>,
    ranges: &mut HashSet<crate::Rect>,
    hyperlinks: &mut Vec<Hyperlink>,
) -> Result<()> {
    if hyperlinks.len() >= MAX_HYPERLINKS {
        return Err(invalid(format!(
            "XLSX worksheet exceeds the {MAX_HYPERLINKS} hyperlink safety limit"
        )));
    }
    let mut reference = None;
    let mut relationship_id = None;
    let mut location = None;
    let mut display = None;
    let mut tooltip = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid(format!("invalid XLSX hyperlink attribute: {error}")))?;
        validate_raw_name(attribute.key.as_ref(), "hyperlink attribute name")?;
        let name = attribute.key.as_ref();
        if name == b"xmlns" || name.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolve_namespace(&namespace)?;
        let local = decode(local.as_ref(), "hyperlink attribute local name")?;
        if attribute.value.as_ref().len() > super::model::MAX_HYPERLINK_TEXT_BYTES {
            return Err(invalid(format!(
                "XLSX hyperlink attribute exceeds the {} byte safety limit",
                super::model::MAX_HYPERLINK_TEXT_BYTES
            )));
        }
        if attribute
            .value
            .as_ref()
            .iter()
            .any(|byte| matches!(byte, b'\t' | b'\n' | b'\r'))
        {
            return Err(invalid(
                "XLSX hyperlink attribute contains non-stable XML whitespace",
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid XLSX hyperlink attribute value: {error}")))?
            .into_owned();
        validate_text(&value, "hyperlink attribute")?;
        match (namespace.as_deref(), local.as_str()) {
            (None, "ref") => reference = Some(value),
            (None, "location") => location = Some(value),
            (None, "display") => display = Some(value),
            (None, "tooltip") => tooltip = Some(value),
            (Some(namespace), "id") if namespace == conformance.relationship_namespace() => {
                if value.len() > MAX_RELATIONSHIP_ID_BYTES {
                    return Err(invalid(format!(
                        "XLSX hyperlink relationship ID exceeds the {MAX_RELATIONSHIP_ID_BYTES} byte limit"
                    )));
                }
                relationship_id = Some(value);
            },
            _ => {
                return Err(invalid(format!(
                    "XLSX hyperlink has unsupported attribute '{local}'"
                )));
            },
        }
    }
    let reference = reference.ok_or_else(|| invalid("XLSX hyperlink is missing ref"))?;
    let reference = HyperlinkReference::parse(&reference)?;
    if !ranges.insert(reference.range()) {
        return Err(invalid(format!(
            "XLSX worksheet has duplicate hyperlink reference '{}'",
            reference.as_str()
        )));
    }
    let external_target = relationship_id
        .map(|relationship_id| {
            let relationship = relationships
                .get(&relationship_id)
                .ok_or_else(|| invalid("XLSX hyperlink relationship is missing"))?;
            if relationship.reltype() != conformance.hyperlink_relationship_type() {
                return Err(invalid("XLSX relationship is not a hyperlink relationship"));
            }
            if !relationship.is_external() {
                return Err(invalid("XLSX hyperlink relationship must be external"));
            }
            used_relationships.insert(relationship_id);
            Ok(relationship.target_ref().to_owned())
        })
        .transpose()?;
    hyperlinks.push(Hyperlink::from_parts(
        reference,
        location,
        external_target,
        display,
        tooltip,
    )?);
    Ok(())
}

fn is_main_namespace(namespace: &str) -> bool {
    matches!(namespace, TRANSITIONAL_MAIN | STRICT_MAIN)
}

fn decode_element_local(element: &BytesStart<'_>) -> Result<String> {
    validate_raw_name(element.name().as_ref(), "element name")?;
    decode(element.local_name().as_ref(), "element local name")
}

fn validate_raw_name(bytes: &[u8], label: &str) -> Result<()> {
    if bytes.len() > MAX_XML_NAME_BYTES {
        return Err(invalid(format!(
            "XLSX hyperlink {label} exceeds the {MAX_XML_NAME_BYTES} byte name limit"
        )));
    }
    Ok(())
}

fn resolve_namespace(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => Ok(Some(decode(value, "namespace URI")?)),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => {
            if prefix.len() > MAX_XML_NAME_BYTES {
                return Err(invalid("XLSX hyperlink XML prefix exceeds the name limit"));
            }
            Err(invalid(format!(
                "XLSX hyperlink XML has unbound prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            )))
        },
    }
}

fn decode(bytes: &[u8], label: &str) -> Result<String> {
    if bytes.len() > MAX_XML_NAME_BYTES {
        return Err(invalid(format!(
            "XLSX hyperlink {label} exceeds the {MAX_XML_NAME_BYTES} byte name limit"
        )));
    }
    String::from_utf8(bytes.to_vec())
        .map_err(|_error| invalid(format!("XLSX hyperlink {label} is not valid UTF-8")))
}

fn relationship_ids(xml: &[u8]) -> Result<Vec<Option<String>>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = false;
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut in_container = false;
    let mut ids = Vec::new();
    let mut root_namespace = None;
    let mut open_names = Vec::<Box<[u8]>>::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid XLSX worksheet hyperlink XML: {error}")))?;
        match event {
            Event::Start(element) => {
                let local = decode_element_local(&element)?;
                let resolved = resolve_namespace(&namespace)?;
                if depth == 0 {
                    root_namespace = resolved;
                } else if depth == 1 && local == "hyperlinks" {
                    if resolved
                        .as_deref()
                        .is_none_or(|value| !matches!(value, TRANSITIONAL_MAIN | STRICT_MAIN))
                    {
                        return Err(invalid(
                            "XLSX hyperlink container uses an unsupported namespace",
                        ));
                    }
                    in_container = true;
                } else if in_container && depth == 2 {
                    let conformance = match root_namespace.as_deref() {
                        Some(TRANSITIONAL_MAIN) => Conformance::Transitional,
                        Some(STRICT_MAIN) => Conformance::Strict,
                        _ => return Err(invalid("XLSX worksheet uses an unsupported namespace")),
                    };
                    require_name(resolved.as_deref(), &local, conformance, "hyperlink")?;
                    ids.push(read_relationship_id(&reader, &element, conformance)?);
                }
                open_names.push(element.name().as_ref().to_vec().into_boxed_slice());
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XLSX worksheet hyperlink XML depth overflows usize"))?;
            },
            Event::Empty(element) => {
                let local = decode_element_local(&element)?;
                let resolved = resolve_namespace(&namespace)?;
                if depth == 0 {
                    root_namespace = resolved;
                } else if depth == 1 && local == "hyperlinks" {
                    in_container = false;
                } else if in_container && depth == 2 {
                    let conformance = match root_namespace.as_deref() {
                        Some(TRANSITIONAL_MAIN) => Conformance::Transitional,
                        Some(STRICT_MAIN) => Conformance::Strict,
                        _ => return Err(invalid("XLSX worksheet uses an unsupported namespace")),
                    };
                    require_name(resolved.as_deref(), &local, conformance, "hyperlink")?;
                    ids.push(read_relationship_id(&reader, &element, conformance)?);
                }
            },
            Event::End(element) => {
                let expected = open_names
                    .pop()
                    .ok_or_else(|| invalid("XLSX worksheet hyperlink XML is not balanced"))?;
                if expected.as_ref() != element.name().as_ref() {
                    return Err(invalid("XLSX worksheet start and end element names differ"));
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    invalid("XLSX worksheet hyperlink XML depth underflows usize")
                })?;
                if depth == 1 && in_container {
                    in_container = false;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if root_namespace
        .as_deref()
        .is_none_or(|value| !matches!(value, TRANSITIONAL_MAIN | STRICT_MAIN))
    {
        return Err(invalid("XLSX worksheet uses an unsupported namespace"));
    }
    Ok(ids)
}

fn validate_exclusive_relationship_references(
    xml: &[u8],
    hyperlink_ids: &[Option<String>],
) -> Result<()> {
    let mut counts = HashMap::<&str, usize>::new();
    counts
        .try_reserve(hyperlink_ids.len())
        .map_err(|source| allocation("worksheet hyperlink relationship reference index", source))?;
    for relationship_id in hyperlink_ids.iter().filter_map(Option::as_deref) {
        if counts.insert(relationship_id, 0).is_some() {
            return Err(invalid(
                "XLSX hyperlink relationship ID is shared by multiple hyperlinks",
            ));
        }
    }
    if counts.is_empty() {
        return Ok(());
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = false;
    reader.config_mut().trim_text(false);
    loop {
        let (_namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid XLSX worksheet hyperlink XML: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(|error| {
                        invalid(format!(
                            "invalid XLSX worksheet relationship attribute: {error}"
                        ))
                    })?;
                    validate_raw_name(
                        attribute.key.as_ref(),
                        "worksheet relationship attribute name",
                    )?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    let namespace = resolve_namespace(&namespace)?;
                    if local.as_ref() != b"id"
                        || namespace.as_deref().is_none_or(|namespace| {
                            !matches!(namespace, TRANSITIONAL_REL | STRICT_REL)
                        })
                    {
                        continue;
                    }
                    if attribute.value.as_ref().len() > MAX_RELATIONSHIP_ID_BYTES {
                        return Err(invalid(format!(
                            "XLSX worksheet relationship ID exceeds the {MAX_RELATIONSHIP_ID_BYTES} byte limit"
                        )));
                    }
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                        .map_err(|error| {
                            invalid(format!("invalid XLSX worksheet relationship ID: {error}"))
                        })?;
                    if let Some(count) = counts.get_mut(value.as_ref()) {
                        *count = count.checked_add(1).ok_or_else(|| {
                            invalid("XLSX worksheet relationship reference count overflows usize")
                        })?;
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if counts.values().any(|count| *count != 1) {
        return Err(invalid(
            "XLSX hyperlink relationship ID is also used by unsupported worksheet markup",
        ));
    }
    Ok(())
}

fn read_relationship_id(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    conformance: Conformance,
) -> Result<Option<String>> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid(format!("invalid XLSX hyperlink attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolve_namespace(&namespace)?;
        if namespace.as_deref() == Some(conformance.relationship_namespace())
            && decode(local.as_ref(), "hyperlink relationship attribute")? == "id"
        {
            return Ok(Some(
                attribute
                    .normalized_value(XmlVersion::Implicit1_0)
                    .map_err(|error| {
                        invalid(format!("invalid XLSX hyperlink relationship ID: {error}"))
                    })?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

/// Rewrite only the direct worksheet hyperlink owner. Other worksheet bytes,
/// including unsupported producer markup, remain byte-identical. Unsupported
/// content inside the hyperlink owner is refused rather than approximated.
pub(crate) fn rewrite_hyperlinks(
    xml: &[u8],
    values: &[Hyperlink],
    relationship_ids: &[Option<&str>],
) -> Result<Vec<u8>> {
    rewrite_hyperlinks_checked(xml, values, relationship_ids, || Ok(()))
}

pub(crate) fn rewrite_hyperlinks_checked<F>(
    xml: &[u8],
    values: &[Hyperlink],
    relationship_ids: &[Option<&str>],
    mut check: F,
) -> Result<Vec<u8>>
where
    F: FnMut() -> Result<()>,
{
    check()?;
    if values.len() != relationship_ids.len() {
        return Err(invalid(
            "XLSX hyperlink rewrite values and relationships differ",
        ));
    }
    if values.len() > MAX_HYPERLINKS {
        return Err(invalid(format!(
            "XLSX worksheet exceeds the {MAX_HYPERLINKS} hyperlink safety limit"
        )));
    }
    let layout = scan_rewrite_layout_checked(xml, &mut check)?;
    if values.is_empty() && layout.container_span.is_none() {
        return Ok(xml.to_vec());
    }
    let replacement = write_hyperlinks(
        layout.element_prefix.as_deref().unwrap_or(""),
        layout.relationship_namespace,
        values,
        relationship_ids,
        &mut check,
    )?;
    if let Some(span) = layout.container_span {
        return replace_span(xml, span, &replacement);
    }
    if layout.root_empty {
        let root = &xml[layout.root_start..layout.root_end];
        let close = root
            .iter()
            .rposition(|byte| *byte == b'/')
            .ok_or_else(|| invalid("XLSX worksheet root is malformed"))?;
        let mut output = Vec::new();
        reserve_output(
            &mut output,
            xml.len()
                .saturating_add(replacement.len())
                .saturating_add(32),
        )?;
        output.extend_from_slice(&xml[..layout.root_start]);
        output.extend_from_slice(&root[..close]);
        output.push(b'>');
        output.extend_from_slice(&replacement);
        output.extend_from_slice(b"</");
        output.extend_from_slice(&layout.root_name);
        output.push(b'>');
        output.extend_from_slice(&xml[layout.root_end..]);
        return Ok(output);
    }
    if layout.unknown_direct_child {
        return Err(invalid(
            "XLSX hyperlink insertion refuses unknown direct worksheet children",
        ));
    }
    let insertion = layout
        .successor_start
        .or(layout.root_close_start)
        .ok_or_else(|| invalid("XLSX worksheet root is malformed"))?;
    if insertion < layout.root_end
        || layout
            .root_close_start
            .is_some_and(|root_close| insertion > root_close)
    {
        return Err(invalid(format!(
            "XLSX hyperlink insertion offset {insertion} is outside worksheet root bounds {}..{}",
            layout.root_end,
            layout.root_close_start.unwrap_or(xml.len())
        )));
    }
    let mut output = Vec::new();
    reserve_output(&mut output, xml.len().saturating_add(replacement.len()))?;
    output.extend_from_slice(&xml[..insertion]);
    output.extend_from_slice(&replacement);
    output.extend_from_slice(&xml[insertion..]);
    check()?;
    Ok(output)
}

pub(crate) fn relationship_type(xml: &[u8]) -> Result<&'static str> {
    match scan_rewrite_layout(xml)?.relationship_namespace {
        TRANSITIONAL_REL => Ok(rt::HYPERLINK),
        STRICT_REL => Ok(rt::STRICT_HYPERLINK),
        _ => Err(invalid(
            "XLSX worksheet relationship namespace is unsupported",
        )),
    }
}

#[derive(Debug)]
struct RewriteLayout {
    root_start: usize,
    root_end: usize,
    root_empty: bool,
    root_name: Box<[u8]>,
    root_close_start: Option<usize>,
    container_span: Option<(usize, usize)>,
    successor_start: Option<usize>,
    element_prefix: Option<String>,
    relationship_namespace: &'static str,
    unknown_direct_child: bool,
}

fn scan_rewrite_layout(xml: &[u8]) -> Result<RewriteLayout> {
    scan_rewrite_layout_checked(xml, &mut || Ok(()))
}

fn scan_rewrite_layout_checked<F>(xml: &[u8], check: &mut F) -> Result<RewriteLayout>
where
    F: FnMut() -> Result<()>,
{
    if xml.len() > MAX_WORKSHEET_XML_BYTES {
        return Err(invalid(format!(
            "XLSX worksheet XML exceeds the {MAX_WORKSHEET_XML_BYTES} byte safety limit"
        )));
    }
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().check_end_names = false;
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut root = None;
    let mut root_close_start = None;
    let mut container_start = None;
    let mut container_end = None;
    let mut successor_start = None;
    let mut stack = Vec::<(Box<[u8]>, usize, usize)>::new();
    let mut container_depth = None;
    let mut unknown_direct_child = false;
    let mut event_count = 0usize;
    loop {
        if event_count % 256 == 0 {
            check()?;
        }
        event_count = event_count
            .checked_add(1)
            .ok_or_else(|| invalid("XLSX worksheet XML event count overflows usize"))?;
        if event_count > MAX_XML_EVENTS {
            return Err(invalid(format!(
                "XLSX worksheet exceeds the {MAX_XML_EVENTS} XML event safety limit"
            )));
        }
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid XLSX worksheet hyperlink XML: {error}")))?;
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("XLSX worksheet XML position overflows usize"))?;
        match event {
            Event::Start(element) => {
                let start = event_start(xml, end)?;
                let name = element.name().as_ref().to_vec().into_boxed_slice();
                let local = local_name(&name);
                if depth == 0 {
                    let namespace = root_namespace(&element)?;
                    let relationship_namespace = match namespace {
                        TRANSITIONAL_MAIN => TRANSITIONAL_REL,
                        STRICT_MAIN => STRICT_REL,
                        _ => return Err(invalid("XLSX worksheet uses an unsupported namespace")),
                    };
                    if local != "worksheet" {
                        return Err(invalid("XLSX hyperlink rewrite requires a worksheet root"));
                    }
                    root = Some((
                        start,
                        end,
                        false,
                        name.clone(),
                        element_prefix(&name),
                        relationship_namespace,
                    ));
                } else if depth == 1 && local == "hyperlinks" {
                    if container_start.is_some() {
                        return Err(invalid(
                            "XLSX worksheet has duplicate hyperlinks containers",
                        ));
                    }
                    container_start = Some(start);
                    container_depth = Some(depth + 1);
                } else if depth == 1 {
                    if !is_known_worksheet_child(local) || !same_element_prefix(&name, &stack[0].0)
                    {
                        unknown_direct_child = true;
                    }
                    if successor_start.is_none() && is_hyperlink_successor(local) {
                        successor_start = Some(start);
                    }
                } else if container_depth == Some(depth) && local != "hyperlink" {
                    return Err(invalid(
                        "XLSX hyperlink owner contains unsupported child markup",
                    ));
                } else if container_depth.is_some() && depth > container_depth.unwrap_or(0) {
                    return Err(invalid(
                        "XLSX hyperlink owner contains unsupported child markup",
                    ));
                }
                stack.push((name, start, end));
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XLSX worksheet XML depth overflows usize"))?;
            },
            Event::Empty(element) => {
                let start = event_start(xml, end)?;
                let name = element.name().as_ref().to_vec();
                let local = local_name(&name);
                if depth == 0 {
                    let namespace = root_namespace(&element)?;
                    let relationship_namespace = match namespace {
                        TRANSITIONAL_MAIN => TRANSITIONAL_REL,
                        STRICT_MAIN => STRICT_REL,
                        _ => return Err(invalid("XLSX worksheet uses an unsupported namespace")),
                    };
                    if local != "worksheet" {
                        return Err(invalid("XLSX hyperlink rewrite requires a worksheet root"));
                    }
                    root = Some((
                        start,
                        end,
                        true,
                        name.clone().into_boxed_slice(),
                        element_prefix(&name),
                        relationship_namespace,
                    ));
                } else if depth == 1 && local == "hyperlinks" {
                    if container_start.is_some() {
                        return Err(invalid(
                            "XLSX worksheet has duplicate hyperlinks containers",
                        ));
                    }
                    container_start = Some(start);
                    container_end = Some(end);
                } else if depth == 1 {
                    if !is_known_worksheet_child(local) || !same_element_prefix(&name, &stack[0].0)
                    {
                        unknown_direct_child = true;
                    }
                    if successor_start.is_none() && is_hyperlink_successor(local) {
                        successor_start = Some(start);
                    }
                } else if container_depth == Some(depth) && local != "hyperlink" {
                    return Err(invalid(
                        "XLSX hyperlink owner contains unsupported child markup",
                    ));
                } else if container_depth.is_some() && depth > container_depth.unwrap_or(0) {
                    return Err(invalid(
                        "XLSX hyperlink owner contains unsupported child markup",
                    ));
                }
            },
            Event::End(element) => {
                let (name, _, _) = stack
                    .pop()
                    .ok_or_else(|| invalid("XLSX worksheet XML is not balanced"))?;
                if name.as_ref() != element.name().as_ref() {
                    return Err(invalid("XLSX worksheet start and end element names differ"));
                }
                if depth == 1 {
                    root_close_start = Some(event_start(xml, end)?);
                }
                if container_depth == Some(depth) {
                    container_end = Some(end);
                    container_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("XLSX worksheet XML depth underflows usize"))?;
            },
            Event::Text(text)
                if container_depth.is_some()
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid(
                    "XLSX hyperlink owner contains unsupported text content",
                ));
            },
            Event::Comment(_) | Event::PI(_) | Event::CData(_) | Event::GeneralRef(_)
                if container_depth.is_some() =>
            {
                return Err(invalid(
                    "XLSX hyperlink owner contains unsupported opaque markup",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let (root_start, root_end, root_empty, root_name, element_prefix, relationship_namespace) =
        root.ok_or_else(|| invalid("XLSX worksheet root is missing"))?;
    if !root_empty && (depth != 0 || !stack.is_empty()) {
        return Err(invalid("XLSX worksheet hyperlink XML is not balanced"));
    }
    check()?;
    Ok(RewriteLayout {
        root_start,
        root_end,
        root_empty,
        root_name,
        root_close_start,
        container_span: container_start.zip(container_end),
        successor_start,
        element_prefix,
        relationship_namespace,
        unknown_direct_child,
    })
}

fn is_known_worksheet_child(local: &str) -> bool {
    matches!(
        local,
        "sheetPr"
            | "dimension"
            | "sheetViews"
            | "sheetFormatPr"
            | "cols"
            | "sheetData"
            | "sheetCalcPr"
            | "sheetProtection"
            | "protectedRanges"
            | "scenarios"
            | "autoFilter"
            | "sortState"
            | "dataConsolidate"
            | "customSheetViews"
            | "mergeCells"
            | "phoneticPr"
            | "conditionalFormatting"
            | "dataValidations"
            | "hyperlinks"
            | "printOptions"
            | "pageMargins"
            | "pageSetup"
            | "headerFooter"
            | "rowBreaks"
            | "colBreaks"
            | "customProperties"
            | "cellWatches"
            | "ignoredErrors"
            | "smartTags"
            | "drawing"
            | "legacyDrawing"
            | "legacyDrawingHF"
            | "picture"
            | "oleObjects"
            | "controls"
            | "webPublishItems"
            | "tableParts"
            | "extLst"
    )
}

fn same_element_prefix(left: &[u8], right: &[u8]) -> bool {
    element_name_prefix(left) == element_name_prefix(right)
}

fn element_name_prefix(name: &[u8]) -> &[u8] {
    name.iter()
        .position(|byte| *byte == b':')
        .map_or(&[], |colon| &name[..colon])
}

fn is_hyperlink_successor(local: &str) -> bool {
    matches!(
        local,
        "printOptions"
            | "pageMargins"
            | "pageSetup"
            | "headerFooter"
            | "rowBreaks"
            | "colBreaks"
            | "customProperties"
            | "cellWatches"
            | "ignoredErrors"
            | "smartTags"
            | "drawing"
            | "legacyDrawing"
            | "legacyDrawingHF"
            | "picture"
            | "oleObjects"
            | "controls"
            | "webPublishItems"
            | "tableParts"
            | "extLst"
    )
}

fn write_hyperlinks<F>(
    prefix: &str,
    relationship_namespace: &str,
    values: &[Hyperlink],
    relationship_ids: &[Option<&str>],
    check: &mut F,
) -> Result<Vec<u8>>
where
    F: FnMut() -> Result<()>,
{
    check()?;
    if values.is_empty() {
        return Ok(Vec::new());
    }
    let qualified_container = format!("{prefix}hyperlinks");
    let qualified_link = format!("{prefix}hyperlink");
    let has_external = relationship_ids.iter().any(Option::is_some);
    let mut output = Vec::new();
    output
        .try_reserve(values.len().saturating_mul(128).saturating_add(64))
        .map_err(|source| allocation("XLSX hyperlink XML", source))?;
    output.extend_from_slice(b"<");
    output.extend_from_slice(qualified_container.as_bytes());
    if has_external {
        output.extend_from_slice(b" xmlns:r=\"");
        output.extend_from_slice(relationship_namespace.as_bytes());
        output.extend_from_slice(b"\"");
    }
    output.extend_from_slice(b">");
    for (index, (value, relationship_id)) in values.iter().zip(relationship_ids).enumerate() {
        if index % 256 == 0 {
            check()?;
        }
        output.extend_from_slice(b"<");
        output.extend_from_slice(qualified_link.as_bytes());
        write_attribute(&mut output, "ref", value.reference().as_str());
        if let Some(location) = value.location() {
            write_attribute(&mut output, "location", location);
        }
        if let Some(display) = value.display() {
            write_attribute(&mut output, "display", display);
        }
        if let Some(tooltip) = value.tooltip() {
            write_attribute(&mut output, "tooltip", tooltip);
        }
        if let Some(relationship_id) = relationship_id {
            write_attribute(&mut output, "r:id", relationship_id);
        }
        output.extend_from_slice(b"/>");
    }
    output.extend_from_slice(b"</");
    output.extend_from_slice(qualified_container.as_bytes());
    output.push(b'>');
    check()?;
    Ok(output)
}

fn write_attribute(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    output.extend_from_slice(litchi_core::xml::escape_xml(value).as_bytes());
    output.push(b'\"');
}

fn replace_span(xml: &[u8], span: (usize, usize), replacement: &[u8]) -> Result<Vec<u8>> {
    let (start, end) = span;
    if start > end || end > xml.len() {
        return Err(invalid("XLSX hyperlink replacement span is invalid"));
    }
    let size = xml
        .len()
        .checked_sub(end - start)
        .and_then(|size| size.checked_add(replacement.len()))
        .ok_or_else(|| invalid("XLSX hyperlink output size overflows usize"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(size)
        .map_err(|source| allocation("XLSX hyperlink output", source))?;
    output.extend_from_slice(&xml[..start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&xml[end..]);
    Ok(output)
}

fn reserve_output(output: &mut Vec<u8>, size: usize) -> Result<()> {
    output
        .try_reserve_exact(size)
        .map_err(|source| allocation("XLSX hyperlink output", source))
}

fn event_start(xml: &[u8], end: usize) -> Result<usize> {
    xml.get(..end)
        .and_then(|bytes| bytes.iter().rposition(|byte| *byte == b'<'))
        .ok_or_else(|| invalid("XLSX worksheet XML event has no start offset"))
}

fn local_name(name: &[u8]) -> &str {
    let name = name.split(|byte| *byte == b':').next_back().unwrap_or(name);
    std::str::from_utf8(name).unwrap_or("")
}

fn element_prefix(name: &[u8]) -> Option<String> {
    let colon = name.iter().position(|byte| *byte == b':')?;
    Some(String::from_utf8_lossy(&name[..=colon]).into_owned())
}

fn root_namespace(element: &BytesStart<'_>) -> Result<&'static str> {
    let element_name = element.name();
    let prefix = element_name
        .as_ref()
        .split(|byte| *byte == b':')
        .next()
        .unwrap_or_default();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid(format!("invalid XLSX worksheet root attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns"
            || (!prefix.is_empty()
                && attribute.key.as_ref().starts_with(b"xmlns:")
                && attribute.key.as_ref().get(6..) == Some(prefix))
        {
            let value = attribute
                .normalized_value(XmlVersion::Implicit1_0)
                .map_err(|error| invalid(format!("invalid XLSX worksheet namespace: {error}")))?;
            return match value.as_ref() {
                TRANSITIONAL_MAIN => Ok(TRANSITIONAL_MAIN),
                STRICT_MAIN => Ok(STRICT_MAIN),
                _ => Err(invalid("XLSX worksheet uses an unsupported namespace")),
            };
        }
    }
    Err(invalid("XLSX worksheet default namespace is missing"))
}

#[cfg(test)]
mod tests {
    use litchi_opc::{Relationships, TargetMode};

    use super::{
        MAX_XML_NAME_BYTES, STRICT_MAIN, STRICT_REL, TRANSITIONAL_MAIN, TRANSITIONAL_REL, parse,
        parse_with_event_limit, parse_with_relationship_ids,
    };
    use litchi_opc::constants::relationship_type as rt;

    fn relationships(reltype: &str, target_mode: TargetMode) -> Relationships {
        let mut relationships = Relationships::new("/xl/worksheets/".to_string());
        relationships
            .try_add_relationship(
                reltype.to_string(),
                "https://127.0.0.1:9/never?q=1#frag".to_string(),
                "rIdHyperlink1".to_string(),
                target_mode,
            )
            .expect("test relationship");
        relationships
    }

    #[test]
    fn transitional_and_strict_links_are_typed_and_inert() {
        for (main, rel_namespace, reltype) in [
            (TRANSITIONAL_MAIN, TRANSITIONAL_REL, rt::HYPERLINK),
            (STRICT_MAIN, STRICT_REL, rt::STRICT_HYPERLINK),
        ] {
            let xml = format!(
                r#"<worksheet xmlns="{main}" xmlns:r="{rel_namespace}"><sheetData/><hyperlinks><hyperlink ref="$A$1" location="Sheet2!A1" display="local"/><hyperlink ref="B2:C3" r:id="rIdHyperlink1" location="anchor" tooltip="tip &amp; more"/></hyperlinks></worksheet>"#
            );
            let values = parse(
                xml.as_bytes(),
                &relationships(reltype, TargetMode::External),
            )
            .expect("typed hyperlinks");
            assert_eq!(values.len(), 2);
            assert_eq!(values[0].reference().as_str(), "$A$1");
            assert_eq!(values[0].location(), Some("Sheet2!A1"));
            assert_eq!(values[0].display(), Some("local"));
            assert_eq!(values[1].reference().range().a1(), "B2:C3");
            assert_eq!(values[1].location(), Some("anchor"));
            assert_eq!(
                values[1].external_target(),
                Some("https://127.0.0.1:9/never?q=1#frag")
            );
            assert_eq!(values[1].tooltip(), Some("tip & more"));
        }
    }

    #[test]
    fn hyperlink_relationship_ids_are_exclusively_owned() {
        let relationships = relationships(rt::HYPERLINK, TargetMode::External);
        let shared = format!(
            r#"<worksheet xmlns="{TRANSITIONAL_MAIN}" xmlns:r="{TRANSITIONAL_REL}"><sheetData/><hyperlinks><hyperlink ref="A1" r:id="rIdHyperlink1"/><hyperlink ref="B2" r:id="rIdHyperlink1"/></hyperlinks></worksheet>"#
        );
        assert!(parse_with_relationship_ids(shared.as_bytes(), &relationships).is_err());

        let opaque = format!(
            r#"<worksheet xmlns="{TRANSITIONAL_MAIN}" xmlns:r="{TRANSITIONAL_REL}"><sheetData/><drawing r:id="rIdHyperlink1"/><hyperlinks><hyperlink ref="A1" r:id="rIdHyperlink1"/></hyperlinks></worksheet>"#
        );
        assert!(parse_with_relationship_ids(opaque.as_bytes(), &relationships).is_err());
    }

    #[test]
    fn malformed_or_opaque_hyperlink_graphs_fail_closed() {
        let empty = Relationships::new("/xl/worksheets/".to_string());
        for body in [
            r#"<hyperlink location="Sheet2!A1"/>"#,
            r#"<hyperlink ref="A0" location="Sheet2!A1"/>"#,
            r#"<hyperlink ref="A1" location="one"/><hyperlink ref="$A$1" location="two"/>"#,
            r#"<hyperlink ref="A1" location="one" vendor="opaque"/>"#,
            r#"<hyperlink ref="A1" location="one"><vendor/></hyperlink>"#,
        ] {
            let xml = format!(
                r#"<worksheet xmlns="{TRANSITIONAL_MAIN}"><sheetData/><hyperlinks>{body}</hyperlinks></worksheet>"#
            );
            assert!(parse(xml.as_bytes(), &empty).is_err(), "accepted {body}");
        }

        let xml = format!(
            r#"<worksheet xmlns="{TRANSITIONAL_MAIN}" xmlns:r="{TRANSITIONAL_REL}"><sheetData/><hyperlinks><hyperlink ref="A1" r:id="rIdHyperlink1"/></hyperlinks></worksheet>"#
        );
        assert!(parse(xml.as_bytes(), &empty).is_err());
        assert!(
            parse(
                xml.as_bytes(),
                &relationships(rt::TABLE, TargetMode::External)
            )
            .is_err()
        );
        assert!(
            parse(
                xml.as_bytes(),
                &relationships(rt::HYPERLINK, TargetMode::Internal)
            )
            .is_err()
        );
    }

    #[test]
    fn orphan_hyperlink_relationship_is_rejected() {
        let xml = format!(r#"<worksheet xmlns="{TRANSITIONAL_MAIN}"><sheetData/></worksheet>"#);
        assert!(
            parse(
                xml.as_bytes(),
                &relationships(rt::HYPERLINK, TargetMode::External)
            )
            .is_err()
        );
    }

    #[test]
    fn malformed_roots_names_and_event_floods_are_rejected() {
        let empty = Relationships::new("/xl/worksheets/".to_string());
        for xml in [
            format!(
                r#"<worksheet xmlns="{TRANSITIONAL_MAIN}"/><worksheet xmlns="{TRANSITIONAL_MAIN}"/>"#
            ),
            format!(r#"<worksheet xmlns="{TRANSITIONAL_MAIN}"/>trailing"#),
            format!(r#"<worksheet xmlns="{TRANSITIONAL_MAIN}">junk<sheetData/></worksheet>"#),
        ] {
            assert!(parse(xml.as_bytes(), &empty).is_err(), "accepted {xml}");
        }

        let long_name = "x".repeat(MAX_XML_NAME_BYTES + 1);
        let xml = format!(r#"<worksheet xmlns="{TRANSITIONAL_MAIN}"><{long_name}/></worksheet>"#);
        assert!(parse(xml.as_bytes(), &empty).is_err());

        let long_prefix = "p".repeat(MAX_XML_NAME_BYTES + 1);
        let xml = format!(
            r#"<{long_prefix}:worksheet xmlns:{long_prefix}="{TRANSITIONAL_MAIN}"></{long_prefix}:worksheet>"#
        );
        assert!(parse(xml.as_bytes(), &empty).is_err());

        let long_end = "e".repeat(MAX_XML_NAME_BYTES + 1);
        let xml = format!(r#"<worksheet xmlns="{TRANSITIONAL_MAIN}"></{long_end}:worksheet>"#);
        assert!(parse(xml.as_bytes(), &empty).is_err());

        let xml =
            format!(r#"<worksheet xmlns="{TRANSITIONAL_MAIN}"><sheetData/><cols/></worksheet>"#);
        assert!(parse_with_event_limit(xml.as_bytes(), &empty, 2).is_err());
    }
}
