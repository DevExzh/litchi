//! Passive dependency policy for removing a workbook model.

use super::codec::{Node, parse_document};
use super::{
    CONNECTIONS_CONTENT_TYPE, CONNECTIONS_RELATIONSHIP_TYPE, Definition, MAX_STRING_BYTES,
    MAX_XML_BYTES, SML, STRICT_CONNECTIONS_RELATIONSHIP_TYPE, STRICT_SML, X15, invalid, xml_error,
};
use crate::error::{Error, Result};
use litchi_core::xml::ReaderOrigin;
use litchi_opc::{OpcPackage, OwnedElementEdit, OwnedElementUpdate, OwnedXmlPart, PackURI};
use quick_xml::{events::Event, reader::NsReader};
use std::collections::{HashMap, HashSet};

const CONNECTION_EXTENSION: &str = "{DE250136-89BD-433C-8126-D09CA5730AF9}";
const QUERY: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml";
const CACHE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
const WORKSHEET: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const PIVOT: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
const REVISION: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml";
const EXTERNAL: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const TABLE: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
const XM: &str = "http://schemas.microsoft.com/office/excel/2006/main";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

pub(super) struct Plan {
    connection_change: Option<ConnectionChange>,
}

enum ConnectionChange {
    Update {
        before: OwnedXmlPart,
        after: OwnedXmlPart,
    },
    Remove {
        part: PackURI,
        workbook: PackURI,
        relationship: String,
        source: OwnedXmlPart,
    },
}

impl Plan {
    pub(super) fn publish(&self, package: &mut OpcPackage) -> Result<()> {
        match &self.connection_change {
            Some(ConnectionChange::Update { before, after }) => {
                package.try_replace_owned_xml_part(before.bytes(), after.clone())?;
            },
            Some(ConnectionChange::Remove {
                part,
                workbook,
                relationship,
                source,
            }) => {
                if package.get_part(part)?.blob() != source.bytes() {
                    return Err(invalid("Connections source changed before owner removal"));
                }
                let before = package.source_relationships(workbook)?;
                let after = before.without_relationship(relationship, before.bytes().len())?;
                package.try_replace_relationships(&before, &after)?;
                if !package.remove_part(part) {
                    return Err(invalid("missing Connections owner during removal"));
                }
            },
            None => {},
        }
        Ok(())
    }
}

pub(super) fn prepare(
    package: &OpcPackage,
    workbook: &PackURI,
    model: &Definition,
) -> Result<Plan> {
    let workbook_part = package.get_part(workbook)?;
    let mut connections = workbook_part.rels().iter().filter(|rel| {
        matches!(
            rel.reltype(),
            CONNECTIONS_RELATIONSHIP_TYPE | STRICT_CONNECTIONS_RELATIONSHIP_TYPE
        )
    });
    let connection = connections.next();
    if connections.next().is_some() {
        return Err(invalid("multiple workbook Connections owners"));
    }
    let mut model_ids = HashSet::new();
    let mut model_names = HashSet::new();
    let mut deleting = HashSet::new();
    let mut source = None;
    let mut removal = None;
    if let Some(connection) = connection {
        if connection.is_external()
            || connection.target_query().is_some()
            || connection.target_fragment().is_some()
        {
            return Err(invalid("invalid Connections target for model removal"));
        }
        let target = connection.target_partname()?;
        let part = package.get_part(&target)?;
        let target = part.partname().clone();
        if part.content_type() != CONNECTIONS_CONTENT_TYPE {
            return Err(invalid("unexpected Connections content type"));
        }
        let values = crate::connections::Connections::parse(part.blob())
            .map_err(|error| invalid(error.to_string()))?;
        model_ids
            .try_reserve(values.connections.len())
            .map_err(|source| litchi_opc::OpcError::Allocation {
                resource: "model connection IDs",
                source,
            })?;
        deleting
            .try_reserve(values.connections.len())
            .map_err(|source| litchi_opc::OpcError::Allocation {
                resource: "removed model connections",
                source,
            })?;
        let root = super::package::parse_connections_document(part.blob())?;
        reject_mce(&root)?;
        let names = model
            .tables
            .iter()
            .map(|table| connection_name(&table.connection))
            .collect::<Result<HashSet<_>>>()?;
        model_names
            .try_reserve(values.connections.len())
            .map_err(|source| litchi_opc::OpcError::Allocation {
                resource: "model connection names",
                source,
            })?;
        let mut nodes = HashMap::new();
        nodes
            .try_reserve(values.connections.len())
            .map_err(|source| litchi_opc::OpcError::Allocation {
                resource: "model connection source owners",
                source,
            })?;
        for node in root
            .children
            .iter()
            .filter(|node| core(node) && node.name == "connection")
        {
            let id = required_u32(node, "id")?;
            if nodes.insert(id, node).is_some() {
                return Err(invalid("duplicate source connection ID"));
            }
        }
        for connection in &values.connections {
            let node = nodes
                .get(&connection.id)
                .copied()
                .ok_or_else(|| invalid("missing source connection owner"))?;
            let extensions: Vec<_> = node
                .children
                .iter()
                .filter(|node| core(node) && node.name == "extLst")
                .flat_map(|node| &node.children)
                .filter(|node| {
                    core(node)
                        && node.name == "ext"
                        && attribute(node, "uri") == Some(CONNECTION_EXTENSION)
                })
                .flat_map(|node| &node.children)
                .filter(|node| node.namespace == X15 && node.name == "connection")
                .collect();
            if extensions.len() > 1 {
                return Err(invalid("multiple extended connection owners"));
            }
            let extension = extensions.first().copied();
            let is_model = extension
                .map(|node| boolean(node, "model"))
                .transpose()?
                .unwrap_or(false);
            let auto_delete = extension
                .map(|node| boolean(node, "autoDelete"))
                .transpose()?
                .unwrap_or(false);
            let used_by_addin = extension
                .map(|node| boolean(node, "usedByAddin"))
                .transpose()?
                .unwrap_or(false);
            if is_model
                && (connection.connection_type != Some(5)
                    || extension.and_then(|node| attribute(node, "id")) != Some(""))
            {
                return Err(invalid("invalid model connection identity"));
            }
            let name = connection
                .name
                .as_deref()
                .map(connection_name)
                .transpose()?;
            if is_model || name.as_ref().is_some_and(|name| names.contains(name)) {
                if let Some(name) = name {
                    model_names.insert(name);
                }
                model_ids.insert(connection.id);
                if auto_delete && !used_by_addin {
                    deleting.insert(connection.id);
                }
            }
        }
        let xml = package.source_xml_part(&target)?;
        if deleting.len() == values.connections.len() {
            validate_connection_owner_removal(
                package,
                workbook_part.partname(),
                &target,
                connection.r_id(),
                &root,
            )?;
            removal = Some(ConnectionChange::Remove {
                part: target,
                workbook: workbook_part.partname().clone(),
                relationship: connection.r_id().to_owned(),
                source: xml,
            });
        } else {
            source = Some(xml);
        }
    }
    // Consumers can be unreferenced physical parts; examine their modeled
    // content types instead of relying only on incoming OPC traversal.
    let mut bytes = workbook_part.blob().len();
    for part in package.iter_parts().filter(|part| {
        matches!(
            part.content_type(),
            QUERY | CACHE | WORKSHEET | TABLE | PIVOT | REVISION | EXTERNAL
        )
    }) {
        // Only a modeled consumer's payload is read, so only consumers are
        // decoded (ADR 0030).
        let part = package.get_part(part.partname())?;
        bytes = bytes
            .checked_add(part.blob().len())
            .filter(|bytes| *bytes <= 256 * 1024 * 1024)
            .ok_or_else(|| invalid("model removal dependency XML limit exceeded"))?;
        check_formulas(part.blob())?;
        if matches!(part.content_type(), WORKSHEET | TABLE | REVISION | EXTERNAL) {
            continue;
        }
        let root = parse_document(part.blob())?;
        reject_mce(&root)?;
        if !core(&root) {
            return Err(invalid("unexpected model consumer namespace"));
        }
        check_connection_names(&root, &model_names)?;
        if part.content_type() == PIVOT {
            if root.name != "pivotTableDefinition" {
                return Err(invalid("invalid Pivot Table root"));
            }
            continue;
        }
        let id = if part.content_type() == QUERY {
            if root.name != "queryTable" {
                return Err(invalid("invalid Query Table root"));
            }
            Some(required_u32(&root, "connectionId")?)
        } else {
            if root.name != "pivotCacheDefinition" {
                return Err(invalid("invalid Pivot Cache root"));
            }
            let sources: Vec<_> = root
                .children
                .iter()
                .filter(|node| core(node) && node.name == "cacheSource")
                .collect();
            if sources.len() != 1 {
                return Err(invalid("ambiguous Pivot Cache source"));
            }
            attribute(sources[0], "connectionId")
                .map(|_| required_u32(sources[0], "connectionId"))
                .transpose()?
        };
        if id.is_some_and(|id| model_ids.contains(&id)) {
            return Err(Error::Unsupported {
                feature: "Data Model removal with dependent query or pivot features",
            });
        }
    }
    check_formulas(workbook_part.blob())?;
    let connection_change = if removal.is_some() {
        removal
    } else if deleting.is_empty() {
        None
    } else {
        let before = source.ok_or_else(|| invalid("missing Connections source"))?;
        let mut reader = NsReader::from_reader(before.bytes());
        let origin = ReaderOrigin::of(before.bytes());
        let mut depth = 0usize;
        let mut updates = Vec::new();
        updates
            .try_reserve_exact(deleting.len())
            .map_err(|source| litchi_opc::OpcError::Allocation {
                resource: "model connection removal spans",
                source,
            })?;
        loop {
            let start = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| invalid("Data Model XML offset exceeds usize"))?;
            let event = reader.read_event().map_err(xml_error)?;
            let end = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| invalid("Data Model XML offset exceeds usize"))?;
            let empty = matches!(&event, Event::Empty(_));
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    let namespace = reader.resolver().resolve_element(element.name()).0;
                    let core = matches!(namespace, quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref() == SML.as_bytes() || ns.as_ref() == STRICT_SML.as_bytes());
                    if depth == 1 && core && element.local_name().as_ref() == b"connection" {
                        for attr in element.attributes().with_checks(true) {
                            let attr = attr.map_err(xml_error)?;
                            if attr.key.as_ref() == b"id" {
                                let value = attr
                                    .decoded_and_normalized_value(
                                        quick_xml::XmlVersion::Implicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(xml_error)?;
                                if value
                                    .parse::<u32>()
                                    .ok()
                                    .is_some_and(|id| deleting.contains(&id))
                                {
                                    updates.push(OwnedElementUpdate {
                                        start_tag: start..end,
                                        edit: OwnedElementEdit::Remove,
                                    });
                                }
                            }
                        }
                    }
                    if !empty {
                        depth += 1;
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("unbalanced Connections XML"))?;
                },
                Event::Eof => break,
                _ => {},
            }
        }
        if updates.len() != deleting.len() {
            return Err(invalid("ambiguous connection removal spans"));
        }
        let after = before.update_elements(&updates, MAX_XML_BYTES)?;
        crate::connections::Connections::parse(after.bytes())
            .map_err(|error| invalid(error.to_string()))?;
        Some(ConnectionChange::Update { before, after })
    };
    Ok(Plan { connection_change })
}

fn validate_connection_owner_removal(
    package: &OpcPackage,
    workbook: &PackURI,
    target: &PackURI,
    relationship_id: &str,
    root: &Node,
) -> Result<()> {
    if root
        .children
        .iter()
        .any(|node| !core(node) || node.name != "connection")
        || root
            .attributes
            .iter()
            .any(|attribute| attribute.namespace != MCE || attribute.name != "Ignorable")
    {
        return Err(Error::Unsupported {
            feature: "Data Model removal with unowned Connections content",
        });
    }
    if !package.get_part(target)?.rels().is_empty() {
        return Err(Error::Unsupported {
            feature: "Data Model removal with outgoing Connections relationships",
        });
    }
    if has_foreign_incoming(package, target, Some((workbook, relationship_id)))? {
        return Err(Error::Unsupported {
            feature: "Data Model removal with foreign Connections references",
        });
    }
    Ok(())
}

pub(super) fn has_foreign_incoming(
    package: &OpcPackage,
    target: &PackURI,
    allowed: Option<(&PackURI, &str)>,
) -> Result<bool> {
    for relationship in package.rels().iter() {
        if !relationship.is_external() && relationship.target_partname()?.is_equivalent_to(target) {
            return Ok(true);
        }
    }
    for owner in package.iter_parts() {
        for relationship in owner.rels().iter() {
            if !relationship.is_external()
                && relationship.target_partname()?.is_equivalent_to(target)
                && allowed != Some((owner.partname(), relationship.r_id()))
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn connection_name(value: &str) -> Result<String> {
    Ok(crate::raw::strings::decode_spreadsheet_text(value)?.to_lowercase())
}

fn check_connection_names(node: &Node, model_names: &HashSet<String>) -> Result<()> {
    // Check expanded names even in unknown extension placement: consumers must
    // not escape dependency refusal through malformed extension wrappers.
    let field = match (node.namespace.as_str(), node.name.as_str()) {
        (X14, "sourceConnection") => Some("name"),
        (X15, "queryTable" | "pivotTableUISettings") => Some("sourceDataName"),
        _ => None,
    };
    if let Some(field) = field {
        if let Some(value) = attribute(node, field) {
            if model_names.contains(&connection_name(value)?)
                || model_names.contains(&value.to_lowercase())
            {
                return Err(Error::Unsupported {
                    feature: "Data Model removal with dependent connection-name features",
                });
            }
        } else if field == "name" {
            return Err(invalid("missing pivot source connection name"));
        }
    }
    if core(node) && node.name == "extLst" {
        let mut seen = [false; 3];
        for extension in node
            .children
            .iter()
            .filter(|node| core(node) && node.name == "ext")
        {
            let slot = match attribute(extension, "uri") {
                Some("{F057638F-6D5F-4E77-A914-E7F072B9BCA8}") => {
                    Some((0, X14, "sourceConnection"))
                },
                Some("{883FBD77-0823-4A55-B5E3-86C4891E6966}") => Some((1, X15, "queryTable")),
                Some("{E67621CE-5B39-4880-91FE-76760E9C1902}") => {
                    Some((2, X15, "pivotTableUISettings"))
                },
                _ => None,
            };
            if let Some((slot, namespace, name)) = slot {
                if seen[slot]
                    || extension
                        .children
                        .iter()
                        .filter(|node| node.namespace == namespace && node.name == name)
                        .count()
                        != 1
                {
                    return Err(invalid("ambiguous connection-name consumer extension"));
                }
                seen[slot] = true;
            }
        }
    }
    for child in &node.children {
        check_connection_names(child, model_names)?;
    }
    Ok(())
}

fn core(node: &Node) -> bool {
    matches!(node.namespace.as_str(), SML | STRICT_SML)
}
fn attribute<'a>(node: &'a Node, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attr| attr.namespace.is_empty() && attr.name == name)
        .map(|attr| attr.value.as_str())
}
fn boolean(node: &Node, name: &str) -> Result<bool> {
    match attribute(node, name).map(|value| value.trim_matches([' ', '\t', '\r', '\n'])) {
        None | Some("0" | "false") => Ok(false),
        Some("1" | "true") => Ok(true),
        _ => Err(invalid("invalid extended connection Boolean")),
    }
}
fn required_u32(node: &Node, name: &str) -> Result<u32> {
    attribute(node, name)
        .and_then(|value| value.parse().ok())
        .ok_or_else(|| invalid("invalid model consumer connection ID"))
}
fn reject_mce(node: &Node) -> Result<()> {
    if node.namespace == MCE
        || node.attributes.iter().any(|attribute| {
            attribute.namespace == MCE
                && matches!(attribute.name.as_str(), "ProcessContent" | "MustUnderstand")
        })
    {
        return Err(Error::Unsupported {
            feature: "Data Model removal with unresolved consumer markup compatibility",
        });
    }
    for child in &node.children {
        reject_mce(child)?;
    }
    Ok(())
}

fn check_formulas(xml: &[u8]) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut formula = None::<String>;
    let mut events = 0usize;
    let mut depth = 0usize;
    loop {
        events += 1;
        if events > 10_000_000 {
            return Err(invalid("model removal formula event limit exceeded"));
        }
        let event = reader.read_event().map_err(xml_error)?;
        let namespace = match &event {
            Event::Start(element) | Event::Empty(element) => {
                reader.resolver().resolve_element(element.name()).0
            },
            _ => quick_xml::name::ResolveResult::Unbound,
        };
        if let Event::Start(element) | Event::Empty(element) = &event {
            if matches!(&namespace, quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref() == MCE.as_bytes())
            {
                return Err(Error::Unsupported {
                    feature: "Data Model removal with unresolved consumer markup compatibility",
                });
            }
            for attribute in element.attributes().with_checks(true) {
                let attribute = attribute.map_err(xml_error)?;
                let element_namespace = reader.resolver().resolve_element(element.name()).0;
                let core = matches!(element_namespace, quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref() == SML.as_bytes() || ns.as_ref() == STRICT_SML.as_bytes());
                if core
                    && matches!(
                        (element.local_name().as_ref(), attribute.key.as_ref()),
                        (b"cfvo", b"val")
                            | (b"cacheField" | b"calculatedItem", b"formula")
                            | (b"definedName", b"refersTo")
                    )
                {
                    if attribute.value.len() > MAX_STRING_BYTES {
                        return Err(invalid(
                            "model removal formula attribute length limit exceeded",
                        ));
                    }
                    let value = attribute
                        .decoded_and_normalized_value(
                            quick_xml::XmlVersion::Implicit1_0,
                            reader.decoder(),
                        )
                        .map_err(xml_error)?;
                    if has_cube_function(&value) {
                        return Err(Error::Unsupported {
                            feature: "Data Model removal with cube formula dependencies",
                        });
                    }
                }
                let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                if matches!(namespace, quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref() == MCE.as_bytes())
                    && matches!(local.as_ref(), b"ProcessContent" | b"MustUnderstand")
                {
                    return Err(Error::Unsupported {
                        feature: "Data Model removal with unresolved consumer markup compatibility",
                    });
                }
            }
        }
        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .filter(|depth| *depth <= super::MAX_DEPTH)
                    .ok_or_else(|| invalid("model dependency XML depth limit exceeded"))?;
                if formula.is_some() {
                    return Err(invalid("nested formula XML"));
                }
                let core = matches!(namespace, quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref() == SML.as_bytes() || ns.as_ref() == STRICT_SML.as_bytes());
                let extension_formula = matches!(namespace, quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref() == XM.as_bytes())
                    && element.local_name().as_ref() == b"f";
                if extension_formula
                    || (core
                        && matches!(
                            element.local_name().as_ref(),
                            b"f" | b"formula"
                                | b"formula1"
                                | b"formula2"
                                | b"definedName"
                                | b"calculatedColumnFormula"
                                | b"totalsRowFormula"
                                | b"oldFormula"
                        ))
                {
                    formula = Some(String::new());
                }
            },
            Event::Text(text) if formula.is_some() => {
                let text = text
                    .xml_content(quick_xml::XmlVersion::Implicit1_0)
                    .map_err(xml_error)?;
                append_formula(&mut formula, &text)?;
            },
            Event::CData(text) if formula.is_some() => {
                let text = text
                    .xml_content(quick_xml::XmlVersion::Implicit1_0)
                    .map_err(xml_error)?;
                append_formula(&mut formula, &text)?;
            },
            Event::Empty(_) if formula.is_some() => return Err(invalid("nested formula XML")),
            Event::DocType(_) => return Err(invalid("DTD in model dependency XML")),
            Event::GeneralRef(reference) if formula.is_some() => {
                let reference = reference.decode().map_err(xml_error)?;
                if reference.len() > 16 {
                    return Err(invalid("invalid formula entity reference"));
                }
                let escaped = format!("&{reference};");
                let text = quick_xml::escape::unescape(&escaped).map_err(xml_error)?;
                append_formula(&mut formula, &text)?;
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced model dependency XML"))?;
                if let Some(formula) = formula.take() {
                    if has_cube_function(&formula) {
                        return Err(Error::Unsupported {
                            feature: "Data Model removal with cube formula dependencies",
                        });
                    }
                }
            },
            Event::Eof => {
                if formula.is_some() || depth != 0 {
                    return Err(invalid("unclosed formula XML"));
                }
                break;
            },
            _ => {},
        }
    }
    Ok(())
}

fn append_formula(formula: &mut Option<String>, text: &str) -> Result<()> {
    if let Some(formula) = formula {
        if formula
            .len()
            .checked_add(text.len())
            .is_none_or(|len| len > MAX_STRING_BYTES)
        {
            return Err(invalid("model removal formula length limit exceeded"));
        }
        formula
            .try_reserve(text.len())
            .map_err(|source| litchi_opc::OpcError::Allocation {
                resource: "model dependency formula",
                source,
            })?;
        formula.push_str(text);
    }
    Ok(())
}

fn has_cube_function(formula: &str) -> bool {
    let bytes = formula.as_bytes();
    let mut at = 0usize;
    while at < bytes.len() {
        if matches!(bytes[at], b'\'' | b'"') {
            let quote = bytes[at];
            at += 1;
            while at < bytes.len() {
                if bytes[at] == quote {
                    at += 1;
                    if bytes.get(at) != Some(&quote) {
                        break;
                    }
                }
                at += 1;
            }
        } else if bytes[at].is_ascii_alphabetic() || bytes[at] == b'_' || bytes[at] >= 0x80 {
            let start = at;
            while at < bytes.len()
                && (bytes[at].is_ascii_alphanumeric()
                    || matches!(bytes[at], b'_' | b'.')
                    || bytes[at] >= 0x80)
            {
                at += 1;
            }
            let mut segments = formula[start..at].rsplit('.');
            let name = segments.next().unwrap_or("");
            let builtin_prefix = segments.all(|prefix| {
                prefix.eq_ignore_ascii_case("_xlfn") || prefix.eq_ignore_ascii_case("_xlws")
            });
            while at < bytes.len() && bytes[at].is_ascii_whitespace() {
                at += 1;
            }
            if builtin_prefix
                && bytes.get(at) == Some(&b'(')
                && [
                    "CUBEVALUE",
                    "CUBEMEMBER",
                    "CUBESET",
                    "CUBESETCOUNT",
                    "CUBERANKEDMEMBER",
                    "CUBEKPIMEMBER",
                    "CUBEMEMBERPROPERTY",
                ]
                .iter()
                .any(|candidate| name.eq_ignore_ascii_case(candidate))
            {
                return true;
            }
        } else {
            at += 1;
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn named_consumers_decode_xstrings_and_reject_duplicate_owners() {
        let names = HashSet::from(["thisworkbookdatamodel".to_owned()]);
        let node = parse_document(br#"<x14:sourceConnection xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="ThisWorkbook_x0044_ataModel"/>"#).unwrap();
        assert!(matches!(
            check_connection_names(&node, &names),
            Err(Error::Unsupported { .. })
        ));
        let node = parse_document(br#"<extLst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><ext uri="{F057638F-6D5F-4E77-A914-E7F072B9BCA8}"><x14:sourceConnection name="ordinary"/></ext><ext uri="{F057638F-6D5F-4E77-A914-E7F072B9BCA8}"><x14:sourceConnection name="ordinary"/></ext></extLst>"#).unwrap();
        assert!(check_connection_names(&node, &names).is_err());
    }

    #[test]
    fn x14_validation_wrappers_leave_xm_formula_ownership_to_the_child() {
        for wrapper in ["formula1", "formula2"] {
            for (formula, dependent) in [("SUM(A1)", false), ("CUBEVALUE(A1)", true)] {
                let xml = format!(
                    r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><extLst><ext uri="{{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}}"><x14:dataValidations><x14:dataValidation><x14:{wrapper}><xm:f>{formula}</xm:f></x14:{wrapper}></x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>"#
                );
                let result = check_formulas(xml.as_bytes());
                if dependent {
                    assert!(matches!(result, Err(Error::Unsupported { .. })));
                } else {
                    result.unwrap();
                }
            }
        }
    }

    #[test]
    fn cube_dependency_scan_handles_strings_names_and_xml_entities() {
        for formula in [
            r#"CUBEVALUE("model", "measure")"#,
            "_xlfn.cubeset(A1)",
            "IF(A1,CUBEMEMBER(A2),0)",
        ] {
            assert!(has_cube_function(formula));
        }
        for formula in [
            r#""CUBEVALUE(model)""#,
            "'CUBEVALUE(foo)'!A1",
            "My.CUBEVALUE(A1)",
            "名CUBEVALUE(A1)",
            "SUM(A1:A2)",
        ] {
            assert!(!has_cube_function(formula));
        }
        let xml = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f>C&#85;BEVALUE(A1)</f></worksheet>"#;
        assert!(matches!(
            check_formulas(xml),
            Err(Error::Unsupported { .. })
        ));
        let xml = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f><![CDATA[CUBEMEMBER(A1)]]></f></worksheet>"#;
        assert!(matches!(
            check_formulas(xml),
            Err(Error::Unsupported { .. })
        ));
        for xml in [
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><extLst><xm:f>CUBEVALUE(A1)</xm:f></extLst></worksheet>"#.as_slice(),
            br#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><tableColumns><tableColumn><calculatedColumnFormula>CUBEVALUE(A1)</calculatedColumnFormula></tableColumn></tableColumns></table>"#.as_slice(),
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent/></worksheet>"#.as_slice(),
        ] {
            assert!(matches!(check_formulas(xml), Err(Error::Unsupported { .. })));
        }
        assert!(check_formulas(br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><f><other/></f></worksheet>"#).is_err());
    }
}
