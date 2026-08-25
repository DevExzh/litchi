use std::collections::HashSet;

use litchi_opc::Relationships;
use litchi_opc::constants::relationship_type as rt;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{Hyperlink, HyperlinkReference, validate_text};
use crate::error::{Result, invalid};

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
                } else if container_depth.is_some() {
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
        return Err(invalid(
            "XLSX hyperlink projection requires a worksheet root",
        ));
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

#[cfg(test)]
mod tests {
    use litchi_opc::{Relationships, TargetMode};

    use super::{
        MAX_XML_NAME_BYTES, STRICT_MAIN, STRICT_REL, TRANSITIONAL_MAIN, TRANSITIONAL_REL, parse,
        parse_with_event_limit,
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
