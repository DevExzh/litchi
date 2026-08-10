//! Bounded XML codec for the inline MS-XLDM workbook descriptor.

use std::collections::{BTreeMap, HashMap, HashSet};

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::Result;

use super::model::{Definition, OpaqueXml, Relationship, Table};
use super::{
    MAX_DEPTH, MAX_EXTENSION_BYTES, MAX_NODES, MAX_RELATIONSHIPS, MAX_REWRITE_BYTES,
    MAX_STRING_BYTES, MAX_TABLES, MAX_TOTAL_STRING_BYTES, MAX_XML_BYTES, SML, STRICT_SML, X15,
    invalid, limit, xml_error,
};

#[derive(Clone)]
pub(crate) struct Attribute {
    pub namespace: String,
    pub name: String,
    pub value: String,
}

#[derive(Clone)]
pub(crate) struct Node {
    pub namespace: String,
    pub name: String,
    pub attributes: Vec<Attribute>,
    pub children: Vec<Node>,
    pub text: String,
}

/// Parse an inline `x15:dataModel` descriptor.
pub fn parse_data_model(xml: &[u8]) -> Result<Definition> {
    let root = parse_document(xml)?;
    parse_data_model_node(&root)
}

/// Deterministically serialize an inline `x15:dataModel` descriptor.
pub fn write_data_model(value: &Definition) -> Result<Vec<u8>> {
    validate_definition(value, false)?;
    let mut output = Vec::new();
    output.extend_from_slice(b"<x15:dataModel xmlns:x15=\"");
    escape(&mut output, X15);
    output.push(b'"');
    if value.min_version_load != 5 {
        attr(
            &mut output,
            "minVersionLoad",
            &value.min_version_load.to_string(),
        );
    }
    if value.tables.is_empty() && value.relationships.is_empty() && value.extension_list.is_none() {
        output.extend_from_slice(b"/>");
        return Ok(output);
    }
    output.push(b'>');
    if !value.tables.is_empty() {
        output.extend_from_slice(b"<x15:modelTables>");
        for table in &value.tables {
            output.extend_from_slice(b"<x15:modelTable");
            attr(&mut output, "id", &table.id);
            attr(&mut output, "name", &table.name);
            attr(&mut output, "connection", &table.connection);
            output.extend_from_slice(b"/>");
        }
        output.extend_from_slice(b"</x15:modelTables>");
    }
    if !value.relationships.is_empty() {
        output.extend_from_slice(b"<x15:modelRelationships>");
        for relationship in &value.relationships {
            output.extend_from_slice(b"<x15:modelRelationship");
            attr(&mut output, "fromTable", &relationship.from_table);
            attr(&mut output, "fromColumn", &relationship.from_column);
            attr(&mut output, "toTable", &relationship.to_table);
            attr(&mut output, "toColumn", &relationship.to_column);
            output.extend_from_slice(b"/>");
        }
        output.extend_from_slice(b"</x15:modelRelationships>");
    }
    if let Some(extension) = &value.extension_list {
        output.extend_from_slice(&extension.xml);
    }
    output.extend_from_slice(b"</x15:dataModel>");
    if output.len() > MAX_XML_BYTES {
        return Err(limit("serialized descriptor bytes"));
    }
    Ok(output)
}

pub(crate) fn parse_data_model_node(root: &Node) -> Result<Definition> {
    require(root, X15, "dataModel")?;
    no_attributes(root, &[("", "minVersionLoad")])?;
    whitespace(root)?;
    let min_version_load = optional(root, "", "minVersionLoad")
        .map(|value| {
            value
                .parse::<u8>()
                .map_err(|_source| invalid("minVersionLoad must be an unsigned byte"))
        })
        .transpose()?
        .unwrap_or(5);
    let mut tables = Vec::new();
    let mut relationships = Vec::new();
    let mut extension_list = None;
    let mut stage = 0u8;
    for child in &root.children {
        match child.name.as_str() {
            "modelTables" if child.namespace == X15 && stage == 0 => {
                stage = 1;
                tables = parse_tables(child)?;
            },
            "modelRelationships" if child.namespace == X15 && stage <= 1 => {
                stage = 2;
                relationships = parse_relationships(child)?;
            },
            "extLst" if child.namespace == X15 && stage <= 2 => {
                stage = 3;
                let xml = serialize_node(child)?;
                if xml.len() > MAX_EXTENSION_BYTES {
                    return Err(limit("extension bytes"));
                }
                extension_list = Some(OpaqueXml { xml });
            },
            _ => return Err(invalid("unexpected or out-of-order dataModel child")),
        }
    }
    let value = Definition {
        min_version_load,
        tables,
        relationships,
        extension_list,
    };
    validate_definition(&value, true)?;
    Ok(value)
}

fn parse_tables(node: &Node) -> Result<Vec<Table>> {
    no_attributes(node, &[])?;
    whitespace(node)?;
    if node.children.is_empty() {
        return Err(invalid("modelTables must contain at least one modelTable"));
    }
    if node.children.len() > MAX_TABLES {
        return Err(limit("table count"));
    }
    node.children
        .iter()
        .map(|child| {
            require(child, X15, "modelTable")?;
            no_attributes(child, &[("", "id"), ("", "name"), ("", "connection")])?;
            leaf(child)?;
            Ok(Table {
                id: required(child, "", "id")?.to_owned(),
                name: required(child, "", "name")?.to_owned(),
                connection: required(child, "", "connection")?.to_owned(),
            })
        })
        .collect()
}

fn parse_relationships(node: &Node) -> Result<Vec<Relationship>> {
    no_attributes(node, &[])?;
    whitespace(node)?;
    if node.children.is_empty() {
        return Err(invalid(
            "modelRelationships must contain at least one modelRelationship",
        ));
    }
    if node.children.len() > MAX_RELATIONSHIPS {
        return Err(limit("relationship count"));
    }
    node.children
        .iter()
        .map(|child| {
            require(child, X15, "modelRelationship")?;
            no_attributes(
                child,
                &[
                    ("", "fromTable"),
                    ("", "fromColumn"),
                    ("", "toTable"),
                    ("", "toColumn"),
                ],
            )?;
            leaf(child)?;
            Ok(Relationship {
                from_table: required(child, "", "fromTable")?.to_owned(),
                from_column: required(child, "", "fromColumn")?.to_owned(),
                to_table: required(child, "", "toTable")?.to_owned(),
                to_column: required(child, "", "toColumn")?.to_owned(),
            })
        })
        .collect()
}

pub(crate) fn validate_definition(
    value: &Definition,
    extension_already_parsed: bool,
) -> Result<()> {
    if value.min_version_load < 5 {
        return Err(invalid("minVersionLoad must be at least 5"));
    }
    if value.tables.len() > MAX_TABLES {
        return Err(limit("table count"));
    }
    if value.relationships.len() > MAX_RELATIONSHIPS {
        return Err(limit("relationship count"));
    }
    let mut ids = HashSet::new();
    let mut names = HashSet::new();
    for table in &value.tables {
        for (field, label) in [
            (&table.id, "table id"),
            (&table.name, "table name"),
            (&table.connection, "connection name"),
        ] {
            bounded_nonempty(field, label)?;
        }
        if !ids.insert(table.id.to_lowercase()) {
            return Err(invalid(format!(
                "duplicate case-insensitive Data Model table id '{}'",
                table.id
            )));
        }
        if !names.insert(table.name.to_lowercase()) {
            return Err(invalid(format!(
                "duplicate case-insensitive Data Model table name '{}'",
                table.name
            )));
        }
    }
    let mut relationships = HashSet::new();
    for relationship in &value.relationships {
        for (field, label) in [
            (&relationship.from_table, "fromTable"),
            (&relationship.from_column, "fromColumn"),
            (&relationship.to_table, "toTable"),
            (&relationship.to_column, "toColumn"),
        ] {
            bounded_nonempty(field, label)?;
        }
        if !names.contains(&relationship.from_table.to_lowercase()) {
            return Err(invalid(format!(
                "relationship references unknown fromTable '{}'",
                relationship.from_table
            )));
        }
        if !names.contains(&relationship.to_table.to_lowercase()) {
            return Err(invalid(format!(
                "relationship references unknown toTable '{}'",
                relationship.to_table
            )));
        }
        let key = (
            relationship.from_table.to_lowercase(),
            relationship.from_column.to_lowercase(),
            relationship.to_table.to_lowercase(),
            relationship.to_column.to_lowercase(),
        );
        if !relationships.insert(key) {
            return Err(invalid(
                "duplicate case-insensitive Data Model relationship",
            ));
        }
    }
    if let Some(extension) = &value.extension_list {
        if extension.xml.len() > MAX_EXTENSION_BYTES {
            return Err(limit("extension bytes"));
        }
        if !extension_already_parsed {
            let root = parse_document(&extension.xml)?;
            require(&root, X15, "extLst")?;
        }
    }
    Ok(())
}

pub(crate) fn workbook_definition(root: &Node) -> Result<(&str, Option<Definition>)> {
    if root.name != "workbook" || !(root.namespace == SML || root.namespace == STRICT_SML) {
        return Err(invalid("expected SpreadsheetML workbook root"));
    }
    let core = root.namespace.as_str();
    let lists: Vec<_> = root
        .children
        .iter()
        .filter(|child| child.namespace == core && child.name == "extLst")
        .collect();
    if lists.len() > 1 {
        return Err(invalid("workbook has multiple direct extLst elements"));
    }
    let mut found = None;
    if let Some(list) = lists.first() {
        for extension in &list.children {
            if extension.namespace == core
                && extension.name == "ext"
                && optional(extension, "", "uri") == Some(super::DATA_MODEL_EXTENSION_URI)
            {
                if found.is_some() {
                    return Err(invalid("workbook has multiple dataModel extensions"));
                }
                no_attributes(extension, &[("", "uri")])?;
                whitespace(extension)?;
                if extension.children.len() != 1 {
                    return Err(invalid(
                        "dataModel extension must contain exactly one dataModel element",
                    ));
                }
                found = Some(parse_data_model_node(&extension.children[0])?);
            }
        }
    }
    Ok((core, found))
}

pub(crate) fn insert_extension(xml: &[u8], core: &str, fragment: &[u8]) -> Result<Vec<u8>> {
    let new_size = xml
        .len()
        .checked_add(fragment.len())
        .and_then(|size| size.checked_add(core.len() + 64))
        .ok_or_else(|| limit("rewrite bytes"))?;
    if new_size > MAX_REWRITE_BYTES {
        return Err(limit("rewrite bytes"));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut open_ext = false;
    let mut empty_ext = None;
    let mut root_close = None;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_source| limit("rewrite position"))?;
        let event = reader.read_event().map_err(xml_error)?;
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_source| limit("rewrite position"))?;
        match event {
            Event::Start(element) => {
                let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
                if depth == 1 && namespace == core && element.local_name().as_ref() == b"extLst" {
                    open_ext = true;
                }
                depth += 1;
            },
            Event::Empty(element) => {
                let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
                if depth == 1 && namespace == core && element.local_name().as_ref() == b"extLst" {
                    empty_ext = Some((start, end, element.name().as_ref().to_vec()));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected workbook closing element"));
                }
                if depth == 2 && open_ext && element.local_name().as_ref() == b"extLst" {
                    let mut output = Vec::with_capacity(new_size);
                    output.extend_from_slice(&xml[..start]);
                    output.extend_from_slice(fragment);
                    output.extend_from_slice(&xml[start..]);
                    return Ok(output);
                }
                if depth == 1 {
                    root_close = Some(start);
                }
                depth -= 1;
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if let Some((start, end, qname)) = empty_ext {
        let raw = &xml[start..end];
        let close = memchr::memmem::rfind(raw, b"/>")
            .ok_or_else(|| invalid("invalid empty workbook extLst"))?;
        let mut output = Vec::with_capacity(new_size);
        output.extend_from_slice(&xml[..start + close]);
        output.push(b'>');
        output.extend_from_slice(fragment);
        output.extend_from_slice(b"</");
        output.extend_from_slice(&qname);
        output.push(b'>');
        output.extend_from_slice(&xml[end..]);
        return Ok(output);
    }
    let close = root_close.ok_or_else(|| invalid("missing workbook closing element"))?;
    let mut output = Vec::with_capacity(new_size);
    output.extend_from_slice(&xml[..close]);
    output.extend_from_slice(b"<extLst xmlns=\"");
    escape(&mut output, core);
    output.extend_from_slice(b"\">");
    output.extend_from_slice(fragment);
    output.extend_from_slice(b"</extLst>");
    output.extend_from_slice(&xml[close..]);
    Ok(output)
}

pub(crate) fn parse_document(xml: &[u8]) -> Result<Node> {
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("XML bytes"));
    }
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut reader = NsReader::from_reader(xml);
    let mut stack = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        let event = reader.read_event().map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(limit("XML structure"));
                }
                let empty = matches!(&event, Event::Empty(_));
                let node = make_node(&reader, element, reader.decoder(), &mut strings)?;
                if empty {
                    attach(node, &mut stack, &mut root)?;
                } else {
                    stack.push(node);
                }
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                attach(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                add_strings(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value = reference
                    .resolve_char_ref()
                    .map_err(xml_error)?
                    .map(|value| value.to_string())
                    .or_else(|| match name.as_ref() {
                        "amp" => Some("&".into()),
                        "lt" => Some("<".into()),
                        "gt" => Some(">".into()),
                        "apos" => Some("'".into()),
                        "quot" => Some("\"".into()),
                        _ => None,
                    })
                    .ok_or_else(|| invalid("custom XML entity is rejected"))?;
                add_strings(&mut strings, value.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else {
                    return Err(invalid("entity outside XML root"));
                }
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected in Data Model XML")),
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated XML"));
    }
    root.ok_or_else(|| invalid("missing XML root"))
}

fn make_node(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    strings: &mut usize,
) -> Result<Node> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    add_strings(strings, namespace.len() + name.len())?;
    let mut attributes = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let qname = item.key.as_ref();
        if qname == b"xmlns" || qname.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let name = std::str::from_utf8(local.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        add_strings(strings, namespace.len() + name.len() + value.len())?;
        if attributes
            .iter()
            .any(|attribute: &Attribute| attribute.namespace == namespace && attribute.name == name)
        {
            return Err(invalid("duplicate expanded XML attribute"));
        }
        attributes.push(Attribute {
            namespace,
            name,
            value,
        });
    }
    Ok(Node {
        namespace,
        name,
        attributes,
        children: Vec::new(),
        text: String::new(),
    })
}

fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

fn serialize_node(node: &Node) -> Result<Vec<u8>> {
    let mut namespaces = BTreeMap::new();
    collect_namespaces(node, &mut namespaces);
    let mut prefixes = HashMap::new();
    let mut next = 0usize;
    for namespace in namespaces.keys() {
        let prefix = match namespace.as_str() {
            X15 => "x15".into(),
            SML | STRICT_SML => "x".into(),
            _ => {
                let prefix = format!("n{next}");
                next += 1;
                prefix
            },
        };
        prefixes.insert(namespace.clone(), prefix);
    }
    let mut output = Vec::new();
    write_node(node, &prefixes, true, &mut output);
    Ok(output)
}

fn collect_namespaces(node: &Node, output: &mut BTreeMap<String, ()>) {
    if !node.namespace.is_empty() {
        output.insert(node.namespace.clone(), ());
    }
    for attribute in &node.attributes {
        if !attribute.namespace.is_empty() {
            output.insert(attribute.namespace.clone(), ());
        }
    }
    for child in &node.children {
        collect_namespaces(child, output);
    }
}

fn write_node(node: &Node, prefixes: &HashMap<String, String>, root: bool, output: &mut Vec<u8>) {
    output.push(b'<');
    qname(output, &node.namespace, &node.name, prefixes);
    if root {
        let mut values: Vec<_> = prefixes.iter().collect();
        values.sort_by(|a, b| a.1.cmp(b.1));
        for (namespace, prefix) in values {
            output.extend_from_slice(b" xmlns:");
            output.extend_from_slice(prefix.as_bytes());
            output.extend_from_slice(b"=\"");
            escape(output, namespace);
            output.push(b'"');
        }
    }
    for attribute in &node.attributes {
        output.push(b' ');
        qname(output, &attribute.namespace, &attribute.name, prefixes);
        output.extend_from_slice(b"=\"");
        escape(output, &attribute.value);
        output.push(b'"');
    }
    if node.children.is_empty() && node.text.is_empty() {
        output.extend_from_slice(b"/>");
        return;
    }
    output.push(b'>');
    escape_text(output, &node.text);
    for child in &node.children {
        write_node(child, prefixes, false, output);
    }
    output.extend_from_slice(b"</");
    qname(output, &node.namespace, &node.name, prefixes);
    output.push(b'>');
}

fn qname(output: &mut Vec<u8>, namespace: &str, name: &str, prefixes: &HashMap<String, String>) {
    if !namespace.is_empty() {
        output.extend_from_slice(prefixes[namespace].as_bytes());
        output.push(b':');
    }
    output.extend_from_slice(name.as_bytes());
}

fn require(node: &Node, namespace: &str, name: &str) -> Result<()> {
    if node.namespace == namespace && node.name == name {
        Ok(())
    } else {
        Err(invalid(format!("expected {{{namespace}}}{name}")))
    }
}

fn no_attributes(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    for attribute in &node.attributes {
        if !allowed
            .iter()
            .any(|(namespace, name)| attribute.namespace == *namespace && attribute.name == *name)
        {
            return Err(invalid(format!(
                "unexpected attribute '{}' on {}",
                attribute.name, node.name
            )));
        }
    }
    Ok(())
}

fn required<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a str> {
    optional(node, namespace, name)
        .ok_or_else(|| invalid(format!("missing required {name} attribute")))
}

fn optional<'a>(node: &'a Node, namespace: &str, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace == namespace && attribute.name == name)
        .map(|attribute| attribute.value.as_str())
}

fn whitespace(node: &Node) -> Result<()> {
    if node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("unexpected text in {}", node.name)))
    }
}

fn leaf(node: &Node) -> Result<()> {
    whitespace(node)?;
    if node.children.is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("{} cannot have children", node.name)))
    }
}

fn bounded_nonempty(value: &str, label: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("{label} cannot be empty")));
    }
    if value.len() > MAX_STRING_BYTES {
        return Err(limit(label));
    }
    Ok(())
}

fn add_strings(total: &mut usize, size: usize) -> Result<()> {
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("XML string bytes"))?;
    if *total > MAX_TOTAL_STRING_BYTES {
        Err(limit("XML string bytes"))
    } else {
        Ok(())
    }
}

fn resolved(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn attr(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escape(output, value);
    output.push(b'"');
}

fn escape(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}

fn escape_text(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}
