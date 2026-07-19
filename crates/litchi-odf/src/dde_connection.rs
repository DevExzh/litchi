//! Inert, bounded OpenDocument text DDE connection declarations.

use crate::{OdfVariableBody, OdfVariableHeaderFooter, OdfVariablePart, OdfVariableScope};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_CONNECTIONS: usize = 65_536;
const MAX_REFERENCES: usize = 262_144;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

/// A named DDE source declaration. It is retained but never contacted.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OdfDdeConnectionDeclaration {
    pub part: OdfVariablePart,
    pub scope: OdfVariableScope,
    pub name: String,
    pub application: String,
    pub topic: String,
    pub item: String,
    pub automatic_update: Option<bool>,
}

impl OdfDdeConnectionDeclaration {
    /// ODF defaults `office:automatic-update` to `false`.
    pub fn effective_automatic_update(&self) -> bool {
        self.automatic_update.unwrap_or(false)
    }
}

/// One `text:dde-connection` occurrence referring to a declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OdfDdeConnectionUse {
    pub part: OdfVariablePart,
    pub scope: OdfVariableScope,
    pub connection_name: String,
}

#[derive(Default)]
pub(crate) struct ParsedDdeConnections {
    pub declarations: Vec<OdfDdeConnectionDeclaration>,
    pub uses: Vec<OdfDdeConnectionUse>,
}

#[derive(Clone)]
struct Frame {
    namespace: Option<String>,
    local: String,
    master_page_name: Option<String>,
}

struct ActiveGroup {
    depth: usize,
    part: OdfVariablePart,
    scope: OdfVariableScope,
}

struct PendingElement {
    depth: usize,
}

type Attributes = HashMap<(String, String), String>;

pub(crate) fn parse_dde_connection_parts(
    parts: &[(&str, OdfVariablePart)],
) -> Result<ParsedDdeConnections> {
    let total = parts.iter().try_fold(0usize, |total, (xml, _)| {
        total
            .checked_add(xml.len())
            .ok_or_else(|| make_error("DDE connection XML size overflow"))
    })?;
    if total > MAX_XML_BYTES {
        return invalid("DDE connection XML exceeds 64 MiB");
    }

    let mut parsed = ParsedDdeConnections::default();
    let mut names = HashSet::<String>::new();
    let mut containers = HashSet::<(OdfVariablePart, OdfVariableScope)>::new();
    let mut aggregate = 0usize;
    for (xml, part) in parts {
        parse_part(
            xml,
            *part,
            &mut parsed,
            &mut names,
            &mut containers,
            &mut aggregate,
        )?;
    }
    for usage in &parsed.uses {
        if !names.contains(&usage.connection_name) {
            return invalid(format!(
                "DDE connection '{}' is used without a declaration",
                usage.connection_name
            ));
        }
    }
    Ok(parsed)
}

fn parse_part(
    xml: &str,
    part: OdfVariablePart,
    parsed: &mut ParsedDdeConnections,
    names: &mut HashSet<String>,
    containers: &mut HashSet<(OdfVariablePart, OdfVariableScope)>,
    aggregate: &mut usize,
) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut stack = Vec::<Frame>::new();
    let mut active: Option<ActiveGroup> = None;
    let mut pending_declaration: Option<PendingElement> = None;
    let mut pending_use: Option<PendingElement> = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| make_error(format!("invalid DDE connection XML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                if pending_declaration.is_some() || pending_use.is_some() {
                    return invalid("DDE connection elements cannot contain elements");
                }
                let namespace = namespace_uri(&namespace)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace.as_deref(), &local)?;
                if let Some(group) = active.as_ref() {
                    if namespace.as_deref() != Some(TEXT)
                        || local != "dde-connection-decl"
                        || depth != group.depth
                    {
                        return invalid(
                            "text:dde-connection-decls may contain only connection declarations",
                        );
                    }
                    let declaration = parse_declaration(
                        &reader,
                        element,
                        group.part,
                        group.scope.clone(),
                        aggregate,
                    )?;
                    add_declaration(declaration, parsed, names)?;
                    pending_declaration = Some(PendingElement { depth: depth + 1 });
                } else if namespace.as_deref() == Some(TEXT) && local == "dde-connection-decls" {
                    start_group(element, part, depth, &stack, containers, &mut active)?;
                } else if namespace.as_deref() == Some(TEXT) && local == "dde-connection" {
                    let usage = parse_use(&reader, element, part, &stack, aggregate)?;
                    add_use(usage, parsed)?;
                    pending_use = Some(PendingElement { depth: depth + 1 });
                }
                let master_page_name =
                    if namespace.as_deref() == Some(STYLE) && local == "master-page" {
                        optional_attribute(&reader, element, STYLE, "name")?
                    } else {
                        None
                    };
                stack.push(Frame {
                    namespace,
                    local,
                    master_page_name,
                });
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| make_error("DDE connection XML depth overflow"))?;
                if depth > MAX_DEPTH {
                    return invalid(format!(
                        "DDE connection XML nesting exceeds {MAX_DEPTH} levels"
                    ));
                }
            },
            Event::Empty(ref element) => {
                if pending_declaration.is_some() || pending_use.is_some() {
                    return invalid("DDE connection elements cannot contain elements");
                }
                let namespace = namespace_uri(&namespace)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace.as_deref(), &local)?;
                if let Some(group) = active.as_ref() {
                    if namespace.as_deref() != Some(TEXT)
                        || local != "dde-connection-decl"
                        || depth != group.depth
                    {
                        return invalid(
                            "text:dde-connection-decls may contain only connection declarations",
                        );
                    }
                    let declaration = parse_declaration(
                        &reader,
                        element,
                        group.part,
                        group.scope.clone(),
                        aggregate,
                    )?;
                    add_declaration(declaration, parsed, names)?;
                } else if namespace.as_deref() == Some(TEXT) && local == "dde-connection-decls" {
                    let mut temporary = None;
                    start_group(element, part, depth, &stack, containers, &mut temporary)?;
                } else if namespace.as_deref() == Some(TEXT) && local == "dde-connection" {
                    let usage = parse_use(&reader, element, part, &stack, aggregate)?;
                    add_use(usage, parsed)?;
                }
            },
            Event::End(_) => {
                if pending_declaration
                    .as_ref()
                    .is_some_and(|pending| pending.depth == depth)
                {
                    pending_declaration = None;
                }
                if pending_use
                    .as_ref()
                    .is_some_and(|pending| pending.depth == depth)
                {
                    pending_use = None;
                }
                if active.as_ref().is_some_and(|group| group.depth == depth) {
                    active = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| make_error("DDE connection XML stack underflow"))?;
                stack
                    .pop()
                    .ok_or_else(|| make_error("DDE connection XML frame stack underflow"))?;
            },
            Event::Text(ref text) => {
                let value = text
                    .decode()
                    .map_err(|error| make_error(format!("invalid DDE connection text: {error}")))?;
                if (pending_declaration.is_some() || pending_use.is_some()) && !value.is_empty() {
                    return invalid("DDE connection elements must be empty");
                }
                if active.is_some() && pending_declaration.is_none() && !value.trim().is_empty() {
                    return invalid(
                        "text:dde-connection-decls may contain only connection declarations",
                    );
                }
            },
            Event::CData(ref value)
                if pending_declaration.is_some() || pending_use.is_some() || active.is_some() =>
            {
                if !value.is_empty() {
                    return invalid("DDE connection elements cannot contain CDATA");
                }
            },
            Event::GeneralRef(_)
                if pending_declaration.is_some() || pending_use.is_some() || active.is_some() =>
            {
                return invalid("DDE connection elements cannot contain entity references");
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTDs and processing instructions are not allowed in DDE XML");
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0
        || !stack.is_empty()
        || active.is_some()
        || pending_declaration.is_some()
        || pending_use.is_some()
    {
        return invalid("incomplete DDE connection XML structure");
    }
    Ok(())
}

fn start_group(
    element: &BytesStart<'_>,
    part: OdfVariablePart,
    depth: usize,
    stack: &[Frame],
    containers: &mut HashSet<(OdfVariablePart, OdfVariableScope)>,
    active: &mut Option<ActiveGroup>,
) -> Result<()> {
    if element.attributes().next().is_some() {
        return invalid("text:dde-connection-decls cannot have attributes");
    }
    let parent = stack
        .last()
        .ok_or_else(|| make_error("misplaced text:dde-connection-decls"))?;
    let scope = body_scope(parent)?;
    if !containers.insert((part, scope.clone())) {
        return invalid("duplicate DDE declaration container in one scope");
    }
    *active = Some(ActiveGroup {
        depth: depth + 1,
        part,
        scope,
    });
    Ok(())
}

fn parse_declaration(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    part: OdfVariablePart,
    scope: OdfVariableScope,
    aggregate: &mut usize,
) -> Result<OdfDdeConnectionDeclaration> {
    let attributes = collect_attributes(reader, element, aggregate)?;
    reject_unexpected(
        &attributes,
        &[
            (OFFICE, "name"),
            (OFFICE, "dde-application"),
            (OFFICE, "dde-topic"),
            (OFFICE, "dde-item"),
            (OFFICE, "automatic-update"),
        ],
    )?;
    let name = required_nonempty(&attributes, OFFICE, "name")?;
    let application = required_nonempty(&attributes, OFFICE, "dde-application")?;
    let topic = required_nonempty(&attributes, OFFICE, "dde-topic")?;
    let item = required_nonempty(&attributes, OFFICE, "dde-item")?;
    let automatic_update = get(&attributes, OFFICE, "automatic-update")
        .map(parse_bool)
        .transpose()?;
    Ok(OdfDdeConnectionDeclaration {
        part,
        scope,
        name,
        application,
        topic,
        item,
        automatic_update,
    })
}

fn parse_use(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    part: OdfVariablePart,
    stack: &[Frame],
    aggregate: &mut usize,
) -> Result<OdfDdeConnectionUse> {
    let attributes = collect_attributes(reader, element, aggregate)?;
    reject_unexpected(&attributes, &[(TEXT, "connection-name")])?;
    Ok(OdfDdeConnectionUse {
        part,
        scope: nearest_scope(stack)?
            .ok_or_else(|| make_error("text:dde-connection occurs outside a document scope"))?,
        connection_name: required_nonempty(&attributes, TEXT, "connection-name")?,
    })
}

fn add_declaration(
    declaration: OdfDdeConnectionDeclaration,
    parsed: &mut ParsedDdeConnections,
    names: &mut HashSet<String>,
) -> Result<()> {
    if parsed.declarations.len() >= MAX_CONNECTIONS {
        return invalid(format!(
            "document exceeds {MAX_CONNECTIONS} DDE connection declarations"
        ));
    }
    if !names.insert(declaration.name.clone()) {
        return invalid(format!(
            "duplicate DDE connection declaration '{}'",
            declaration.name
        ));
    }
    parsed.declarations.push(declaration);
    Ok(())
}

fn add_use(usage: OdfDdeConnectionUse, parsed: &mut ParsedDdeConnections) -> Result<()> {
    if parsed.uses.len() >= MAX_REFERENCES {
        return invalid(format!(
            "document exceeds {MAX_REFERENCES} DDE connection references"
        ));
    }
    parsed.uses.push(usage);
    Ok(())
}

fn collect_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Attributes> {
    let mut attributes = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| make_error(format!("invalid DDE connection attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_uri(&namespace)?.unwrap_or_default();
        let local = decode(local.as_ref(), "attribute name")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| make_error(format!("invalid DDE connection attribute: {error}")))?
            .into_owned();
        if value.len() > MAX_VALUE_BYTES {
            return invalid("DDE connection attribute exceeds 64 KiB");
        }
        *aggregate = aggregate
            .checked_add(value.len())
            .ok_or_else(|| make_error("DDE connection text size overflow"))?;
        if *aggregate > MAX_AGGREGATE_BYTES {
            return invalid("DDE connection metadata exceeds 16 MiB");
        }
        if attributes.insert((namespace, local), value).is_some() {
            return invalid("duplicate expanded DDE connection attribute");
        }
    }
    Ok(attributes)
}

fn optional_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &str,
    local: &str,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| make_error(format!("invalid XML attribute: {error}")))?;
        let (resolved, resolved_local) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_uri(&resolved)?.as_deref() == Some(namespace)
            && resolved_local.as_ref() == local.as_bytes()
        {
            if value.is_some() {
                return invalid(format!("duplicate expanded attribute {local}"));
            }
            value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                    .map_err(|error| make_error(format!("invalid XML attribute value: {error}")))?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

fn body_scope(parent: &Frame) -> Result<OdfVariableScope> {
    if parent.namespace.as_deref() != Some(OFFICE) {
        return invalid("DDE declarations must be direct children of an office body element");
    }
    let body = match parent.local.as_str() {
        "text" => OdfVariableBody::Text,
        "spreadsheet" => OdfVariableBody::Spreadsheet,
        "presentation" => OdfVariableBody::Presentation,
        "drawing" => OdfVariableBody::Drawing,
        "chart" => OdfVariableBody::Chart,
        _ => return invalid("DDE declarations must be direct children of an office body element"),
    };
    Ok(OdfVariableScope::Body(body))
}

fn nearest_scope(stack: &[Frame]) -> Result<Option<OdfVariableScope>> {
    for (index, frame) in stack.iter().enumerate().rev() {
        if frame.namespace.as_deref() == Some(OFFICE)
            && matches!(
                frame.local.as_str(),
                "text" | "spreadsheet" | "presentation" | "drawing" | "chart"
            )
        {
            return body_scope(frame).map(Some);
        }
        if frame.namespace.as_deref() == Some(STYLE) {
            let kind = match frame.local.as_str() {
                "header" => Some(OdfVariableHeaderFooter::Header),
                "header-first" => Some(OdfVariableHeaderFooter::HeaderFirst),
                "header-left" => Some(OdfVariableHeaderFooter::HeaderLeft),
                "footer" => Some(OdfVariableHeaderFooter::Footer),
                "footer-first" => Some(OdfVariableHeaderFooter::FooterFirst),
                "footer-left" => Some(OdfVariableHeaderFooter::FooterLeft),
                _ => None,
            };
            if let Some(kind) = kind {
                let master_page_name = stack[..=index]
                    .iter()
                    .rev()
                    .find_map(|candidate| candidate.master_page_name.clone());
                return Ok(Some(OdfVariableScope::HeaderFooter {
                    kind,
                    master_page_name,
                }));
            }
        }
    }
    Ok(None)
}

fn reject_spoofed_name(namespace: Option<&str>, local: &str) -> Result<()> {
    if matches!(
        local,
        "dde-connection-decls" | "dde-connection-decl" | "dde-connection"
    ) && namespace != Some(TEXT)
    {
        return invalid("DDE connection vocabulary uses the wrong namespace");
    }
    Ok(())
}

fn reject_unexpected(attributes: &Attributes, allowed: &[(&str, &str)]) -> Result<()> {
    for (namespace, local) in attributes.keys() {
        if !allowed.iter().any(|(allowed_namespace, allowed_local)| {
            namespace == allowed_namespace && local == allowed_local
        }) && matches!(namespace.as_str(), OFFICE | TEXT)
        {
            return invalid(format!(
                "unexpected DDE connection attribute {namespace}:{local}"
            ));
        }
    }
    Ok(())
}

fn get<'a>(attributes: &'a Attributes, namespace: &str, local: &str) -> Option<&'a str> {
    attributes
        .get(&(namespace.to_string(), local.to_string()))
        .map(String::as_str)
}

fn required_nonempty(attributes: &Attributes, namespace: &str, local: &str) -> Result<String> {
    match get(attributes, namespace, local) {
        Some(value) if !value.is_empty() => Ok(value.to_string()),
        _ => invalid(format!("DDE connection requires non-empty {local}")),
    }
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("invalid DDE automatic-update value '{value}'")),
    }
}

fn namespace_uri(result: &ResolveResult<'_>) -> Result<Option<String>> {
    match result {
        ResolveResult::Bound(Namespace(value)) => Ok(Some(decode(value, "namespace URI")?)),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => Err(make_error(format!(
            "unbound namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn decode(value: &[u8], description: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| make_error(format!("invalid UTF-8 {description}")))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(make_error(message))
}

fn make_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::variable_declaration::parse_variable_declaration_parts;

    const PREFIX: &str = r#"<o:document-content
        xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
        xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
        xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><o:body><o:text>"#;
    const SUFFIX: &str = "</o:text></o:body></o:document-content>";

    #[test]
    fn retains_inert_declarations_and_references() {
        let xml = format!(
            r#"{PREFIX}<t:dde-connection-decls>
                <t:dde-connection-decl o:name="Prices" o:dde-application="soffice"
                    o:dde-topic="file:///tmp/prices.ods" o:dde-item="Sheet1.A1"
                    o:automatic-update="true"/>
            </t:dde-connection-decls><t:p><t:dde-connection t:connection-name="Prices"/></t:p>{SUFFIX}"#
        );
        let inventory =
            parse_variable_declaration_parts(&[(xml.as_str(), OdfVariablePart::Content)]).unwrap();
        assert_eq!(inventory.dde_connections.len(), 1);
        assert_eq!(inventory.dde_connection_uses.len(), 1);
        let declaration = &inventory.dde_connections[0];
        assert_eq!(declaration.name, "Prices");
        assert_eq!(declaration.topic, "file:///tmp/prices.ods");
        assert!(declaration.effective_automatic_update());
        assert_eq!(inventory.dde_connection_uses[0].connection_name, "Prices");
    }

    #[test]
    fn accepts_declaration_from_another_part_and_default_update() {
        let styles = format!(
            r#"<o:document-styles xmlns:o="{OFFICE}" xmlns:t="{TEXT}"><o:master-styles/></o:document-styles>"#
        );
        let content = format!(
            r#"{PREFIX}<t:dde-connection-decls><t:dde-connection-decl o:name="Feed"
                o:dde-application="app" o:dde-topic="topic" o:dde-item="item"/>
                </t:dde-connection-decls><t:dde-connection t:connection-name="Feed"/>{SUFFIX}"#
        );
        let parsed = parse_variable_declaration_parts(&[
            (styles.as_str(), OdfVariablePart::Styles),
            (content.as_str(), OdfVariablePart::Content),
        ])
        .unwrap();
        assert!(!parsed.dde_connections[0].effective_automatic_update());
    }

    #[test]
    fn rejects_malformed_active_and_unresolved_connections() {
        let bodies = [
            r#"<t:dde-connection-decls><t:dde-connection-decl o:name="x" o:dde-application="a" o:dde-topic="t"/></t:dde-connection-decls>"#,
            r#"<t:dde-connection-decls><t:dde-connection-decl o:name="x" o:dde-application="a" o:dde-topic="t" o:dde-item="i" o:automatic-update="yes"/></t:dde-connection-decls>"#,
            r#"<t:dde-connection-decls><t:dde-connection-decl o:name="x" o:dde-application="a" o:dde-topic="t" o:dde-item="i">payload</t:dde-connection-decl></t:dde-connection-decls>"#,
            r#"<t:dde-connection t:connection-name="missing"/>"#,
            r#"<t:dde-connection-decls><t:p/></t:dde-connection-decls>"#,
        ];
        for body in bodies {
            let xml = format!("{PREFIX}{body}{SUFFIX}");
            assert!(
                parse_variable_declaration_parts(&[(xml.as_str(), OdfVariablePart::Content,)])
                    .is_err(),
                "accepted {body}"
            );
        }
    }
}
