//! Bounded XML codec for spreadsheet named definitions.

use crate::model::names::{Definition, Expression, Range, Scope, Usage};
use crate::model::names::{validate_collection, write_definitions};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashMap;

const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const MAX_CONTENT_BYTES: usize = 16 * 1024 * 1024;
const MAX_DEFINITIONS: usize = 262_144;

#[derive(Debug, Clone)]
struct Span {
    start: usize,
    end: usize,
    scope: Scope,
}

#[derive(Debug, Clone)]
struct Host {
    start: usize,
    end: usize,
    content_start: Option<usize>,
    name: Option<String>,
    qualified_name: String,
}

#[derive(Debug, Clone)]
struct TableFrame {
    name: Option<String>,
}

#[derive(Debug)]
struct Scan {
    definitions: Vec<Definition>,
    containers: Vec<Span>,
    spreadsheet: Option<Host>,
    tables: Vec<Host>,
}

/// Parse and validate the ordered named-definition catalog in content XML.
pub(crate) fn parse(xml: &str) -> Result<Vec<Definition>> {
    let scan = scan(xml)?;
    validate_collection(&scan.definitions)?;
    Ok(scan.definitions)
}

/// Replace the complete named-definition catalog without rewriting unrelated XML.
pub(crate) fn replace(xml: &str, definitions: &[Definition]) -> Result<String> {
    validate_collection(definitions)?;
    let scan = scan(xml)?;

    let groups = groups(definitions);
    let mut first_container = HashMap::with_capacity(scan.containers.len());
    let mut edits = Vec::with_capacity(scan.containers.len() + groups.len());

    for container in &scan.containers {
        let replacement = if first_container
            .insert(container.scope.clone(), container.start)
            .is_none()
        {
            match groups.get(&container.scope) {
                Some(values) => container_xml(values)?,
                None => String::new(),
            }
        } else {
            String::new()
        };
        edits.push(Edit {
            start: container.start,
            end: container.end,
            replacement,
        });
    }

    for (scope, values) in &groups {
        if first_container.contains_key(scope) {
            continue;
        }
        let container = container_xml(values)?;
        let host = match scope {
            Scope::Global => scan.spreadsheet.as_ref(),
            Scope::Sheet(name) => scan
                .tables
                .iter()
                .find(|table| table.name.as_deref() == Some(name.as_str())),
        }
        .ok_or_else(|| missing_host(scope))?;

        if let Some(content_start) = host.content_start {
            edits.push(Edit {
                start: content_start,
                end: content_start,
                replacement: container,
            });
        } else {
            edits.push(Edit {
                start: host.start,
                end: host.end,
                replacement: expand_empty_host(xml, host, &container)?,
            });
        }
    }

    if edits.is_empty() {
        return Ok(xml.to_owned());
    }

    edits.sort_by_key(|edit| (edit.start, edit.end));
    let mut output_len = xml.len();
    for edit in &edits {
        if edit.start > edit.end
            || edit.end > xml.len()
            || !xml.is_char_boundary(edit.start)
            || !xml.is_char_boundary(edit.end)
        {
            return invalid("named-definition XML span is invalid");
        }
        output_len = output_len
            .checked_sub(edit.end - edit.start)
            .and_then(|length| length.checked_add(edit.replacement.len()))
            .ok_or_else(|| invalid_error("named-definition XML size overflow"))?;
        if output_len > MAX_CONTENT_BYTES {
            return invalid("updated ODS content.xml exceeds the mutation limit");
        }
    }

    let mut output = String::with_capacity(output_len);
    let mut cursor = 0usize;
    for edit in edits {
        if edit.start < cursor {
            return invalid("overlapping named-definition XML spans");
        }
        output.push_str(&xml[cursor..edit.start]);
        output.push_str(&edit.replacement);
        cursor = edit.end;
    }
    output.push_str(&xml[cursor..]);
    Ok(output)
}

#[derive(Debug)]
struct Edit {
    start: usize,
    end: usize,
    replacement: String,
}

fn groups(definitions: &[Definition]) -> HashMap<Scope, Vec<&Definition>> {
    let mut groups = HashMap::new();
    for definition in definitions {
        groups
            .entry(definition.scope().clone())
            .or_insert_with(Vec::new)
            .push(definition);
    }
    groups
}

fn container_xml(definitions: &[&Definition]) -> Result<String> {
    let mut output = String::with_capacity(64 + definitions.len() * 96);
    write_definitions(&mut output, definitions.iter().copied());
    let declaration = " xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"";
    let insertion = output
        .find('>')
        .ok_or_else(|| invalid_error("named-definition writer emitted no container"))?;
    output.insert_str(insertion, declaration);
    Ok(output)
}

fn expand_empty_host(xml: &str, host: &Host, child: &str) -> Result<String> {
    let raw = xml
        .get(host.start..host.end)
        .ok_or_else(|| invalid_error("empty named-definition host is outside XML"))?;
    let close = raw
        .rfind("/>")
        .ok_or_else(|| invalid_error("empty named-definition host is not self-closing"))?;
    let mut output = String::with_capacity(raw.len() + child.len() + host.qualified_name.len() + 3);
    output.push_str(&raw[..close]);
    output.push('>');
    output.push_str(child);
    output.push_str("</");
    output.push_str(&host.qualified_name);
    output.push('>');
    Ok(output)
}

fn missing_host(scope: &Scope) -> Error {
    Error::InvalidFormat(match scope {
        Scope::Global => {
            "global named definitions require an office:spreadsheet element".to_string()
        },
        Scope::Sheet(name) => {
            format!("sheet-local named definitions reference missing sheet '{name}'")
        },
    })
}

fn scan(xml: &str) -> Result<Scan> {
    if xml.len() > MAX_CONTENT_BYTES {
        return invalid("ODS content.xml exceeds the mutation limit");
    }

    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut tables = Vec::new();
    let mut table_stack = Vec::<TableFrame>::new();
    let mut containers = Vec::new();
    let mut definitions = Vec::new();
    let mut active: Option<(usize, Scope)> = None;
    let mut spreadsheet = None;

    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid_error("XML position overflow"))?;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid_error(format!("invalid ODS content.xml: {error}")))?;
        let is_table = is_namespace(&namespace, TABLE_NAMESPACE);
        let is_office = is_namespace(&namespace, OFFICE_NAMESPACE);
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid_error("XML position overflow"))?;

        match event {
            Event::Start(element)
                if is_office && element.local_name().as_ref() == b"spreadsheet" =>
            {
                if spreadsheet.is_some() {
                    return invalid("multiple office:spreadsheet elements");
                }
                spreadsheet = Some(host(&element, start, end, Some(end), None)?);
            },
            Event::Empty(element)
                if is_office && element.local_name().as_ref() == b"spreadsheet" =>
            {
                if spreadsheet.is_some() {
                    return invalid("multiple office:spreadsheet elements");
                }
                spreadsheet = Some(host(&element, start, end, None, None)?);
            },
            Event::Start(element) if is_table && element.local_name().as_ref() == b"table" => {
                let name = table_attribute(reader.resolver(), reader.decoder(), &element, b"name")?;
                tables.push(host(&element, start, end, Some(end), name.clone())?);
                table_stack.push(TableFrame { name });
            },
            Event::Empty(element) if is_table && element.local_name().as_ref() == b"table" => {
                let name = table_attribute(reader.resolver(), reader.decoder(), &element, b"name")?;
                tables.push(host(&element, start, end, None, name)?);
            },
            Event::Start(element) if is_table => match element.local_name().as_ref() {
                b"named-expressions" => {
                    if active.is_some() {
                        return invalid("nested table:named-expressions element");
                    }
                    let scope = scope(&table_stack)?;
                    let index = containers.len();
                    containers.push(Span {
                        start,
                        end,
                        scope: scope.clone(),
                    });
                    active = Some((index, scope));
                },
                b"named-range" | b"named-expression" => {
                    if let Some((_, scope)) = &active {
                        if definitions.len() >= MAX_DEFINITIONS {
                            return invalid("named-definition count exceeds the safety limit");
                        }
                        definitions.push(parse_definition(
                            reader.resolver(),
                            reader.decoder(),
                            &element,
                            scope.clone(),
                        )?);
                    }
                },
                _ => {},
            },
            Event::Empty(element) if is_table => match element.local_name().as_ref() {
                b"named-expressions" => {
                    if active.is_some() {
                        return invalid("nested table:named-expressions element");
                    }
                    let scope = scope(&table_stack)?;
                    containers.push(Span { start, end, scope });
                },
                b"named-range" | b"named-expression" => {
                    if let Some((_, scope)) = &active {
                        if definitions.len() >= MAX_DEFINITIONS {
                            return invalid("named-definition count exceeds the safety limit");
                        }
                        definitions.push(parse_definition(
                            reader.resolver(),
                            reader.decoder(),
                            &element,
                            scope.clone(),
                        )?);
                    }
                },
                _ => {},
            },
            Event::End(element) if is_table => match element.local_name().as_ref() {
                b"named-expressions" => {
                    let Some((index, _)) = active.take() else {
                        return invalid("unmatched table:named-expressions end element");
                    };
                    containers[index].end = end;
                },
                b"table" => {
                    if active.is_some() {
                        return invalid("unterminated table:named-expressions element");
                    }
                    table_stack
                        .pop()
                        .ok_or_else(|| invalid_error("table element stack underflow"))?;
                },
                _ => {},
            },
            Event::Eof => break,
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }

    if active.is_some() || !table_stack.is_empty() {
        return invalid("ODS content.xml contains an unclosed named-definition host");
    }
    Ok(Scan {
        definitions,
        containers,
        spreadsheet,
        tables,
    })
}

fn host(
    element: &BytesStart<'_>,
    start: usize,
    end: usize,
    content_start: Option<usize>,
    name: Option<String>,
) -> Result<Host> {
    let qualified_name = String::from_utf8(element.name().as_ref().to_vec())
        .map_err(|_error| invalid_error("ODS element name is not UTF-8"))?;
    Ok(Host {
        start,
        end,
        content_start,
        name,
        qualified_name,
    })
}

fn scope(tables: &[TableFrame]) -> Result<Scope> {
    match tables.last() {
        Some(TableFrame { name: Some(name) }) => Ok(Scope::Sheet(name.clone())),
        Some(TableFrame { name: None }) => {
            invalid("sheet-local named definitions require table:name")
        },
        None => Ok(Scope::Global),
    }
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

fn parse_definition(
    resolver: &NamespaceResolver,
    decoder: Decoder,
    element: &BytesStart<'_>,
    scope: Scope,
) -> Result<Definition> {
    let name = required_table_attribute(resolver, decoder, element, b"name")?;
    let base_cell_address = table_attribute(resolver, decoder, element, b"base-cell-address")?;

    match element.local_name().as_ref() {
        b"named-range" => {
            let cell_range_address =
                required_table_attribute(resolver, decoder, element, b"cell-range-address")?;
            let mut range = Range::new(name, cell_range_address, scope)?;
            range.base_cell_address = base_cell_address;
            if let Some(usable_as) =
                table_attribute(resolver, decoder, element, b"range-usable-as")?
            {
                if usable_as.is_empty() {
                    return invalid("table:range-usable-as must not be empty");
                }
                if usable_as != "none" {
                    for token in usable_as.split_whitespace() {
                        let usage = Usage::parse(token)?;
                        if !range.usable_as.contains(&usage) {
                            range.usable_as.push(usage);
                        }
                    }
                }
            }
            range.validate()?;
            Ok(range.into())
        },
        b"named-expression" => {
            let expression = required_table_attribute(resolver, decoder, element, b"expression")?;
            let namespace_uri = formula_namespace_uri(resolver, &expression)?;
            let mut value = match namespace_uri {
                Some(uri) => Expression::new_with_namespace(name, expression, uri, scope)?,
                None => Expression::new(name, expression, scope)?,
            };
            value.base_cell_address = base_cell_address;
            value.validate()?;
            Ok(value.into())
        },
        _ => invalid("unexpected named-definition element"),
    }
}

fn required_table_attribute(
    resolver: &NamespaceResolver,
    decoder: Decoder,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<String> {
    table_attribute(resolver, decoder, element, local_name)?.ok_or_else(|| {
        invalid_error(format!(
            "{} is missing required table:{} attribute",
            String::from_utf8_lossy(element.local_name().as_ref()),
            String::from_utf8_lossy(local_name)
        ))
    })
}

fn table_attribute(
    resolver: &NamespaceResolver,
    decoder: Decoder,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid_error(format!("invalid XML attribute: {error}")))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if is_namespace(&namespace, TABLE_NAMESPACE) && local.as_ref() == local_name {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| invalid_error(format!("invalid XML attribute value: {error}")))?;
            return Ok(Some(value.into_owned()));
        }
    }
    Ok(None)
}

fn formula_namespace_uri(resolver: &NamespaceResolver, expression: &str) -> Result<Option<String>> {
    let Some((prefix, remainder)) = expression.split_once(':') else {
        return Ok(None);
    };
    if prefix.is_empty() || !remainder.starts_with('=') {
        return Ok(None);
    }
    for (declaration, namespace) in resolver.bindings() {
        if let PrefixDeclaration::Named(candidate) = declaration
            && candidate == prefix.as_bytes()
        {
            return String::from_utf8(namespace.as_ref().to_vec())
                .map(Some)
                .map_err(|_error| {
                    invalid_error(format!(
                        "formula namespace for prefix '{prefix}' is not UTF-8"
                    ))
                });
        }
    }
    Err(Error::InvalidFormat(format!(
        "formula prefix '{prefix}' is not bound to a namespace"
    )))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
