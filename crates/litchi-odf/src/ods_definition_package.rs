//! Atomic, lossless ODS named-definition and database-range package mutation.

use crate::core::OwnedPackage;
use crate::embedded_chart::rebuild_package;
use crate::ods::database_range::{validate_database_range_collection, write_database_range_fragment};
use crate::ods::named_expression::{
    expression_references_name, validate_named_definition_collection, write_named_definition_fragment,
};
use crate::{DatabaseOrientation, DatabaseRange, FormulaNamespace, NamedDefinition, NamedDefinitionScope, NamedRangeUsage};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{Reader, events::{BytesStart, Event}, name::{Namespace, ResolveResult}, reader::NsReader};
use std::collections::HashMap;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TABLE_URI: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const MAX_XML: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct NamedDefinitionUpdate {
    pub name: Option<String>,
    pub base_cell_address: Option<Option<String>>,
    pub cell_range_address: Option<String>,
    pub usable_as: Option<Vec<NamedRangeUsage>>,
    pub expression: Option<String>,
    pub formula_namespace: Option<Option<FormulaNamespace>>,
}

impl NamedDefinitionUpdate {
    fn apply(&self, value: &mut NamedDefinition) -> Result<()> {
        match value {
            NamedDefinition::Range(range) => {
                if self.expression.is_some() || self.formula_namespace.is_some() {
                    return invalid("expression metadata cannot update a named range");
                }
                if let Some(value) = &self.name { range.name = value.clone(); }
                if let Some(value) = &self.base_cell_address { range.base_cell_address = value.clone(); }
                if let Some(value) = &self.cell_range_address { range.cell_range_address = value.clone(); }
                if let Some(value) = &self.usable_as { range.usable_as = value.clone(); }
            },
            NamedDefinition::Expression(expression) => {
                if self.cell_range_address.is_some() || self.usable_as.is_some() {
                    return invalid("range metadata cannot update a named expression");
                }
                if let Some(value) = &self.name { expression.name = value.clone(); }
                if let Some(value) = &self.base_cell_address { expression.base_cell_address = value.clone(); }
                if let Some(value) = &self.expression { expression.expression = value.clone(); }
                if let Some(value) = &self.formula_namespace { expression.formula_namespace = value.clone(); }
            },
        }
        Ok(())
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct DatabaseRangeUpdate {
    pub name: Option<Option<String>>,
    pub is_selection: Option<Option<bool>>,
    pub on_update_keep_styles: Option<Option<bool>>,
    pub on_update_keep_size: Option<Option<bool>>,
    pub has_persistent_data: Option<Option<bool>>,
    pub orientation: Option<Option<DatabaseOrientation>>,
    pub contains_header: Option<Option<bool>>,
    pub display_filter_buttons: Option<Option<bool>>,
    pub target_range_address: Option<String>,
    pub refresh_delay: Option<Option<String>>,
}

impl DatabaseRangeUpdate {
    fn apply(&self, range: &mut DatabaseRange) {
        if let Some(value) = &self.name { range.name = value.clone(); }
        if let Some(value) = self.is_selection { range.is_selection = value; }
        if let Some(value) = self.on_update_keep_styles { range.on_update_keep_styles = value; }
        if let Some(value) = self.on_update_keep_size { range.on_update_keep_size = value; }
        if let Some(value) = self.has_persistent_data { range.has_persistent_data = value; }
        if let Some(value) = self.orientation { range.orientation = value; }
        if let Some(value) = self.contains_header { range.contains_header = value; }
        if let Some(value) = self.display_filter_buttons { range.display_filter_buttons = value; }
        if let Some(value) = &self.target_range_address { range.target_range_address = value.clone(); }
        if let Some(value) = &self.refresh_delay { range.refresh_delay = value.clone(); }
    }
}

#[derive(Clone)]
struct Span { start: usize, end: usize, close: Option<usize>, qname: String }
#[derive(Clone, Copy, PartialEq, Eq)]
enum NamespaceKind { Office, Table, Other }
#[derive(Clone, Copy)]
enum CollectionKind { Named, Database }
struct Scan {
    containers: Vec<(Option<String>, Span)>,
    items: Vec<(Option<String>, Span)>,
    sheets: HashMap<String, Span>,
    insertion: usize,
}

pub(crate) fn add_named(package: &OwnedPackage, xml: &str, current: &[NamedDefinition], value: &NamedDefinition) -> Result<Vec<u8>> {
    let mut proposed = current.to_vec(); proposed.push(value.clone());
    validate_named_definition_collection(&proposed)?;
    let scan = scan(xml, CollectionKind::Named)?;
    let scope = scope_name(value.scope());
    let fragment = write_named_definition_fragment(value)?;
    let updated = if let Some((_, container)) = scan.containers.iter().find(|(name, _)| name.as_deref() == scope.as_deref()) {
        insert_child(xml, container, &fragment)?
    } else {
        let container = format!("<table:named-expressions xmlns:table=\"{TABLE_URI}\">{fragment}</table:named-expressions>");
        let at = if let Some(sheet) = &scope {
            scan.sheets.get(sheet).and_then(|value| value.close)
                .ok_or_else(|| make_error(format!("named definition refers to missing sheet '{sheet}'")))?
        } else { scan.insertion };
        splice(xml, at, at, &container)?
    };
    validate_named_xml(&updated)?; rebuild(package, &updated)
}

pub(crate) fn replace_named(package: &OwnedPackage, xml: &str, current: &[NamedDefinition], name: &str, scope: &NamedDefinitionScope, value: &NamedDefinition) -> Result<Vec<u8>> {
    if value.scope() != scope { return invalid("named-definition replacement cannot change scope"); }
    let (index, mut proposed) = named_copy(current, name, scope)?;
    proposed[index] = value.clone(); validate_named_definition_collection(&proposed)?;
    replace_item(package, xml, CollectionKind::Named, index, &write_named_definition_fragment(value)?, validate_named_xml)
}

pub(crate) fn update_named(package: &OwnedPackage, xml: &str, current: &[NamedDefinition], name: &str, scope: &NamedDefinitionScope, update: &NamedDefinitionUpdate) -> Result<Vec<u8>> {
    let (index, mut proposed) = named_copy(current, name, scope)?;
    update.apply(&mut proposed[index])?; validate_named_definition_collection(&proposed)?;
    let fragment = write_named_definition_fragment(&proposed[index])?;
    let known = match proposed[index] {
        NamedDefinition::Range(_) => &["name", "cell-range-address", "base-cell-address", "range-usable-as"][..],
        NamedDefinition::Expression(_) => &["name", "expression", "base-cell-address"][..],
    };
    update_item(package, xml, CollectionKind::Named, index, &fragment, known, validate_named_xml)
}

pub(crate) fn remove_named(package: &OwnedPackage, xml: &str, current: &[NamedDefinition], name: &str, scope: &NamedDefinitionScope) -> Result<Vec<u8>> {
    let (index, mut proposed) = named_copy(current, name, scope)?; proposed.remove(index);
    for value in &proposed {
        if let NamedDefinition::Expression(expression) = value
            && (scope == &NamedDefinitionScope::Global || &expression.scope == scope)
            && expression_references_name(&expression.expression, name)
        {
            return Err(make_error(format!("named expression '{}' depends on '{name}'", expression.name)));
        }
    }
    validate_named_definition_collection(&proposed)?;
    remove_item(package, xml, CollectionKind::Named, index, validate_named_xml)
}

pub(crate) fn reorder_named(package: &OwnedPackage, xml: &str, current: &[NamedDefinition], scope: &NamedDefinitionScope, from: usize, to: usize) -> Result<Vec<u8>> {
    let indexes = current.iter().enumerate().filter_map(|(index, value)| (value.scope() == scope).then_some(index)).collect::<Vec<_>>();
    let from = *indexes.get(from).ok_or_else(|| bounds(from, indexes.len()))?;
    let to = *indexes.get(to).ok_or_else(|| bounds(to, indexes.len()))?;
    reorder_items(package, xml, CollectionKind::Named, from, to, validate_named_xml)
}

pub(crate) fn add_database(package: &OwnedPackage, xml: &str, current: &[DatabaseRange], value: &DatabaseRange) -> Result<Vec<u8>> {
    let mut proposed = current.to_vec(); proposed.push(value.clone()); validate_database_range_collection(&proposed)?;
    let scan = scan(xml, CollectionKind::Database)?;
    let fragment = write_database_range_fragment(value)?;
    let updated = if let Some((_, container)) = scan.containers.first() { insert_child(xml, container, &fragment)? }
    else {
        let container = format!("<table:database-ranges xmlns:table=\"{TABLE_URI}\">{fragment}</table:database-ranges>");
        splice(xml, scan.insertion, scan.insertion, &container)?
    };
    validate_database_xml(&updated)?; rebuild(package, &updated)
}

pub(crate) fn replace_database(package: &OwnedPackage, xml: &str, current: &[DatabaseRange], index: usize, value: &DatabaseRange) -> Result<Vec<u8>> {
    let mut proposed = current.to_vec(); let len = proposed.len();
    *proposed.get_mut(index).ok_or_else(|| bounds(index, len))? = value.clone();
    validate_database_range_collection(&proposed)?;
    replace_item(package, xml, CollectionKind::Database, index, &write_database_range_fragment(value)?, validate_database_xml)
}

pub(crate) fn update_database(package: &OwnedPackage, xml: &str, current: &[DatabaseRange], index: usize, update: &DatabaseRangeUpdate) -> Result<Vec<u8>> {
    let mut proposed = current.to_vec(); let len = proposed.len();
    update.apply(proposed.get_mut(index).ok_or_else(|| bounds(index, len))?);
    validate_database_range_collection(&proposed)?;
    let known = ["name", "is-selection", "on-update-keep-styles", "on-update-keep-size", "has-persistent-data", "orientation", "contains-header", "display-filter-buttons", "target-range-address", "refresh-delay"];
    update_item(package, xml, CollectionKind::Database, index, &write_database_range_fragment(&proposed[index])?, &known, validate_database_xml)
}

pub(crate) fn remove_database(package: &OwnedPackage, xml: &str, current: &[DatabaseRange], index: usize) -> Result<Vec<u8>> {
    let mut proposed = current.to_vec(); if index >= proposed.len() { return Err(bounds(index, proposed.len())); }
    proposed.remove(index); validate_database_range_collection(&proposed)?;
    remove_item(package, xml, CollectionKind::Database, index, validate_database_xml)
}

pub(crate) fn reorder_database(package: &OwnedPackage, xml: &str, current: &[DatabaseRange], from: usize, to: usize) -> Result<Vec<u8>> {
    if from >= current.len() { return Err(bounds(from, current.len())); }
    if to >= current.len() { return Err(bounds(to, current.len())); }
    reorder_items(package, xml, CollectionKind::Database, from, to, validate_database_xml)
}

fn replace_item(package: &OwnedPackage, xml: &str, kind: CollectionKind, index: usize, fragment: &str, validate: fn(&str) -> Result<()>) -> Result<Vec<u8>> {
    let scan = scan(xml, kind)?; let span = &scan.items.get(index).ok_or_else(|| bounds(index, scan.items.len()))?.1;
    let updated = splice(xml, span.start, span.end, fragment)?; validate(&updated)?; rebuild(package, &updated)
}
fn update_item(package: &OwnedPackage, xml: &str, kind: CollectionKind, index: usize, fragment: &str, known: &[&str], validate: fn(&str) -> Result<()>) -> Result<Vec<u8>> {
    let scan = scan(xml, kind)?; let span = &scan.items.get(index).ok_or_else(|| bounds(index, scan.items.len()))?.1;
    let original = start_tag(&xml[span.start..span.end])?;
    let end = span.start.checked_add(original.len()).ok_or_else(|| make_error("XML position overflow"))?;
    let mut merged = merge_start_tag(original, start_tag(fragment)?, known)?;
    if span.close.is_none() {
        merged.pop();
        merged.push_str("/>");
    }
    let updated = splice(xml, span.start, end, &merged)?; validate(&updated)?; rebuild(package, &updated)
}
fn remove_item(package: &OwnedPackage, xml: &str, kind: CollectionKind, index: usize, validate: fn(&str) -> Result<()>) -> Result<Vec<u8>> {
    let scan = scan(xml, kind)?; let span = &scan.items.get(index).ok_or_else(|| bounds(index, scan.items.len()))?.1;
    let updated = splice(xml, span.start, span.end, "")?; validate(&updated)?; rebuild(package, &updated)
}
fn reorder_items(package: &OwnedPackage, xml: &str, kind: CollectionKind, from: usize, to: usize, validate: fn(&str) -> Result<()>) -> Result<Vec<u8>> {
    if from == to { return Ok(package.as_bytes().to_vec()); }
    let scan = scan(xml, kind)?;
    let left = &scan.items.get(from).ok_or_else(|| bounds(from, scan.items.len()))?.1;
    let right = &scan.items.get(to).ok_or_else(|| bounds(to, scan.items.len()))?.1;
    let updated = swap(xml, left, right)?; validate(&updated)?; rebuild(package, &updated)
}

fn scan(xml: &str, kind: CollectionKind) -> Result<Scan> {
    if xml.len() > MAX_XML { return invalid("ODS mutation XML exceeds 64 MiB"); }
    let (container_local, item_locals): (&[u8], &[&[u8]]) = match kind {
        CollectionKind::Named => (b"named-expressions", &[b"named-range", b"named-expression"]),
        CollectionKind::Database => (b"database-ranges", &[b"database-range"]),
    };
    let mut reader = NsReader::from_str(xml); let mut buffer = Vec::new(); let mut depth = 0usize;
    let mut spreadsheet = None; let mut spreadsheet_close = None; let mut tail = None;
    let mut sheet: Option<(usize, usize, String, String)> = None; let mut sheets = HashMap::new();
    let mut container: Option<(usize, usize, Option<String>, String)> = None; let mut containers = Vec::new();
    let mut item: Option<(usize, usize, Option<String>, String)> = None; let mut items = Vec::new();
    loop {
        let start = position(&reader)?;
        let (resolved, event) = reader.read_resolved_event_into(&mut buffer).map_err(xml_error)?;
        let namespace = classify(&resolved); let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                let local = element.local_name();
                if namespace == NamespaceKind::Office && local.as_ref() == b"spreadsheet" { spreadsheet = Some(depth); }
                else if namespace == NamespaceKind::Table && local.as_ref() == b"table" && spreadsheet.is_some_and(|value| depth == value + 1) && sheet.is_none() {
                    let name = table_attribute(&reader, &element, b"name")?.ok_or_else(|| make_error("sheet lacks table:name"))?;
                    sheet = Some((depth, start, name, qname(element.name().as_ref())?));
                } else if namespace == NamespaceKind::Table && local.as_ref() == container_local
                    && (spreadsheet.is_some_and(|value| depth == value + 1) || matches!(kind, CollectionKind::Named) && sheet.as_ref().is_some_and(|(value, _, _, _)| depth == *value + 1))
                {
                    if container.is_some() { return invalid("nested collection container"); }
                    container = Some((depth, start, sheet.as_ref().map(|(_, _, name, _)| name.clone()), qname(element.name().as_ref())?));
                } else if namespace == NamespaceKind::Table && item_locals.contains(&local.as_ref()) && container.as_ref().is_some_and(|(value, _, _, _)| depth == *value + 1) {
                    item = Some((depth, start, container.as_ref().and_then(|(_, _, scope, _)| scope.clone()), qname(element.name().as_ref())?));
                } else if sheet.is_none() && spreadsheet.is_some_and(|value| depth == value + 1) && namespace == NamespaceKind::Table && is_tail(kind, local.as_ref()) {
                    tail.get_or_insert(start);
                }
                depth = depth.checked_add(1).ok_or_else(|| make_error("XML depth overflow"))?;
                if depth > MAX_DEPTH { return invalid("ODS mutation XML nesting exceeds limit"); }
            },
            Event::Empty(element) => {
                let local = element.local_name();
                if namespace == NamespaceKind::Table && local.as_ref() == container_local
                    && (spreadsheet.is_some_and(|value| depth == value + 1) || matches!(kind, CollectionKind::Named) && sheet.as_ref().is_some_and(|(value, _, _, _)| depth == *value + 1))
                {
                    containers.push((sheet.as_ref().map(|(_, _, name, _)| name.clone()), Span { start, end, close: None, qname: qname(element.name().as_ref())? }));
                } else if namespace == NamespaceKind::Table && item_locals.contains(&local.as_ref()) && container.as_ref().is_some_and(|(value, _, _, _)| depth == *value + 1) {
                    items.push((container.as_ref().and_then(|(_, _, scope, _)| scope.clone()), Span { start, end, close: None, qname: qname(element.name().as_ref())? }));
                } else if sheet.is_none() && spreadsheet.is_some_and(|value| depth == value + 1) && namespace == NamespaceKind::Table && is_tail(kind, local.as_ref()) { tail.get_or_insert(start); }
            },
            Event::End(element) => {
                depth = depth.checked_sub(1).ok_or_else(|| make_error("XML depth underflow"))?;
                if item.as_ref().is_some_and(|(value, _, _, _)| *value == depth) {
                    let (_, item_start, scope, name) = item.take().expect("active item"); items.push((scope, Span { start: item_start, end, close: Some(start), qname: name }));
                } else if container.as_ref().is_some_and(|(value, _, _, _)| *value == depth) {
                    let (_, container_start, scope, name) = container.take().expect("active container"); containers.push((scope, Span { start: container_start, end, close: Some(start), qname: name }));
                } else if sheet.as_ref().is_some_and(|(value, _, _, _)| *value == depth) && namespace == NamespaceKind::Table && element.local_name().as_ref() == b"table" {
                    let (_, sheet_start, name, tag) = sheet.take().expect("active sheet"); sheets.insert(name, Span { start: sheet_start, end, close: Some(start), qname: tag });
                } else if spreadsheet == Some(depth) && namespace == NamespaceKind::Office && element.local_name().as_ref() == b"spreadsheet" { spreadsheet_close = Some(start); spreadsheet = None; }
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODS mutation XML"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    let close = spreadsheet_close.ok_or_else(|| make_error("missing office:spreadsheet"))?;
    let mut scopes = std::collections::HashSet::new();
    for (scope, _) in &containers {
        if !scopes.insert(scope.clone()) { return invalid("duplicate collection container in one scope"); }
    }
    items.sort_by_key(|(_, span)| span.start);
    Ok(Scan { containers, items, sheets, insertion: tail.unwrap_or(close) })
}

fn is_tail(kind: CollectionKind, local: &[u8]) -> bool { match kind {
    CollectionKind::Named => matches!(local, b"database-ranges" | b"data-pilot-tables" | b"consolidation" | b"dde-links"),
    CollectionKind::Database => matches!(local, b"data-pilot-tables" | b"consolidation" | b"dde-links"),
} }
fn scope_name(scope: &NamedDefinitionScope) -> Option<String> { match scope { NamedDefinitionScope::Global => None, NamedDefinitionScope::Sheet(value) => Some(value.clone()) } }
fn named_copy(current: &[NamedDefinition], name: &str, scope: &NamedDefinitionScope) -> Result<(usize, Vec<NamedDefinition>)> {
    let index = current.iter().position(|value| value.name() == name && value.scope() == scope).ok_or_else(|| make_error(format!("named definition '{name}' not found in {scope:?}")))?;
    Ok((index, current.to_vec()))
}
fn validate_named_xml(xml: &str) -> Result<()> { scan(xml, CollectionKind::Named)?; validate_named_definition_collection(&crate::ods::parser::OdsParser::parse_named_definitions(xml)?) }
fn validate_database_xml(xml: &str) -> Result<()> { scan(xml, CollectionKind::Database)?; validate_database_range_collection(&crate::ods::database_range::parse_database_ranges(xml)?) }
fn rebuild(package: &OwnedPackage, xml: &str) -> Result<Vec<u8>> { rebuild_package(package, xml, Vec::new(), Vec::new(), Vec::new(), Vec::new()) }
fn classify(value: &ResolveResult<'_>) -> NamespaceKind { match value { ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_NS => NamespaceKind::Office, ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_NS => NamespaceKind::Table, _ => NamespaceKind::Other } }
fn table_attribute(reader: &NsReader<&[u8]>, element: &BytesStart<'_>, local: &[u8]) -> Result<Option<String>> { for attribute in element.attributes().with_checks(true) { let attribute = attribute.map_err(xml_error)?; if matches!(reader.resolver().resolve_attribute(attribute.key), (ResolveResult::Bound(Namespace(uri)), name) if uri == TABLE_NS && name.as_ref() == local) { return attribute.decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder()).map(|value| Some(value.into_owned())).map_err(xml_error); } } Ok(None) }
fn insert_child(xml: &str, container: &Span, fragment: &str) -> Result<String> { if let Some(close) = container.close { splice(xml, close, close, fragment) } else { let raw = &xml[container.start..container.end]; let slash = raw.rfind("/>").ok_or_else(|| make_error("invalid empty container"))?; splice(xml, container.start, container.end, &format!("{}>{}</{}>", &raw[..slash], fragment, container.qname)) } }
fn swap(xml: &str, a: &Span, b: &Span) -> Result<String> { let (a, b) = if a.start < b.start { (a, b) } else { (b, a) }; if a.end > b.start { return invalid("overlapping mutation spans"); } Ok(format!("{}{}{}{}{}", &xml[..a.start], &xml[b.start..b.end], &xml[a.end..b.start], &xml[a.start..a.end], &xml[b.end..])) }
fn splice(xml: &str, start: usize, end: usize, value: &str) -> Result<String> { if start > end || end > xml.len() || !xml.is_char_boundary(start) || !xml.is_char_boundary(end) { return invalid("invalid mutation span"); } Ok(format!("{}{}{}", &xml[..start], value, &xml[end..])) }
fn start_tag(xml: &str) -> Result<&str> { xml.find('>').map(|end| &xml[..=end]).ok_or_else(|| make_error("unterminated start tag")) }
fn merge_start_tag(original: &str, replacement: &str, known: &[&str]) -> Result<String> { let (name, old) = parse_start(original)?; let (_, mut attrs) = parse_start(replacement)?; let prefix = name.split_once(':').map_or("", |value| value.0); for (attr, value) in old { let (attr_prefix, local) = attr.split_once(':').unwrap_or(("", attr.as_str())); let replace = known.contains(&local) && (attr_prefix == prefix || attr_prefix == "table"); if !replace && !attrs.iter().any(|value| value.0 == attr) { attrs.push((attr, value)); } } let mut output = format!("<{name}"); for (name, value) in attrs { output.push(' '); output.push_str(&name); output.push_str("=\""); output.push_str(&escape_xml(&value)); output.push('"'); } output.push('>'); Ok(output) }
fn parse_start(xml: &str) -> Result<(String, Vec<(String, String)>)> { let mut reader = Reader::from_str(xml); let decoder = reader.decoder(); match reader.read_event().map_err(xml_error)? { Event::Start(element) | Event::Empty(element) => { let name = qname(element.name().as_ref())?; let mut values = Vec::new(); for attribute in element.attributes().with_checks(true) { let attribute = attribute.map_err(xml_error)?; values.push((qname(attribute.key.as_ref())?, attribute.decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder).map_err(xml_error)?.into_owned())); } Ok((name, values)) }, _ => invalid("expected element start") } }
fn position(reader: &NsReader<&[u8]>) -> Result<usize> { usize::try_from(reader.buffer_position()).map_err(|_| make_error("XML position overflow")) }
fn qname(value: &[u8]) -> Result<String> { std::str::from_utf8(value).map(str::to_string).map_err(|_| make_error("invalid qualified name")) }
fn bounds(index: usize, len: usize) -> Error { make_error(format!("index {index} is out of bounds for {len} entries")) }
fn xml_error(error: impl std::fmt::Display) -> Error { make_error(format!("invalid ODS mutation XML: {error}")) }
fn make_error(value: impl Into<String>) -> Error { Error::InvalidFormat(value.into()) }
fn invalid<T>(value: impl Into<String>) -> Result<T> { Err(make_error(value)) }

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        DatabaseFilter, DatabaseSort, DatabaseSortKey, FilterCondition, FilterExpression,
        NamedExpression, NamedRange, SortOrder, Spreadsheet, SpreadsheetBuilder, SubtotalField,
        SubtotalRule, SubtotalRules,
    };
    #[test]
    fn rejects_named_cycles_and_invalid_database_targets() {
        let a = crate::NamedExpression::new("A", "of:=B", NamedDefinitionScope::Global).unwrap().into();
        let b = crate::NamedExpression::new("B", "of:=A", NamedDefinitionScope::Global).unwrap().into();
        assert!(validate_named_definition_collection(&[a, b]).is_err());
        let range = DatabaseRange::new("Sheet1.B2:A1");
        assert!(range.validate().is_err());
    }

    #[test]
    fn generated_package_crud_round_trips_scopes_filters_sorts_and_subtotals() {
        let mut builder = SpreadsheetBuilder::new();
        builder.add_sheet("Sheet1").unwrap();
        let mut spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();

        let global: NamedDefinition = NamedRange::new("GlobalData", "Sheet1.A1:D20", NamedDefinitionScope::Global).unwrap().into();
        let local: NamedDefinition = NamedExpression::new("LocalValue", "of:=1", NamedDefinitionScope::sheet("Sheet1")).unwrap().into();
        spreadsheet.add_named_definition(&global).unwrap();
        spreadsheet.add_named_definition(&local).unwrap();
        assert!(spreadsheet.find_named_definition("GlobalData", &NamedDefinitionScope::Global).is_some());
        spreadsheet.update_named_definition("GlobalData", &NamedDefinitionScope::Global, &NamedDefinitionUpdate { name: Some("RenamedData".to_string()), ..Default::default() }).unwrap();

        let dependent: NamedDefinition = NamedExpression::new("Dependent", "of:=RenamedData", NamedDefinitionScope::Global).unwrap().into();
        spreadsheet.add_named_definition(&dependent).unwrap();
        assert!(spreadsheet.remove_named_definition("RenamedData", &NamedDefinitionScope::Global).is_err());
        spreadsheet.reorder_named_definition(&NamedDefinitionScope::Global, 0, 1).unwrap();

        let mut database = DatabaseRange::new("Sheet1.A1:D20");
        database.name = Some("ImportData".to_string());
        database.display_filter_buttons = Some(true);
        database.refresh_delay = Some("PT5M".to_string());
        database.filter = Some(DatabaseFilter {
            target_range_address: None,
            condition_source: None,
            condition_source_range_address: None,
            display_duplicates: Some(false),
            expression: FilterExpression::Condition(FilterCondition::new(0, "=", "East")),
        });
        database.sort = Some(DatabaseSort {
            keys: vec![DatabaseSortKey { field_number: 1, data_type: Some("text".to_string()), order: Some(SortOrder::Ascending) }],
            ..Default::default()
        });
        database.subtotals = Some(SubtotalRules {
            rules: vec![SubtotalRule { group_by_field_number: 0, fields: vec![SubtotalField { field_number: 3, function: "sum".to_string() }] }],
            ..Default::default()
        });
        spreadsheet.add_database_range(&database).unwrap();
        assert!(spreadsheet.find_database_range("ImportData").is_some());
        spreadsheet.update_database_range(0, &DatabaseRangeUpdate { display_filter_buttons: Some(Some(false)), refresh_delay: Some(None), ..Default::default() }).unwrap();
        assert!(spreadsheet.database_ranges()[0].filter.is_some());
        assert!(spreadsheet.database_ranges()[0].sort.is_some());
        assert!(spreadsheet.database_ranges()[0].subtotals.is_some());
        spreadsheet.remove_database_range(0).unwrap();
    }

    #[test]
    fn update_merge_preserves_foreign_attributes_and_self_closing_shape() {
        let merged = merge_start_tag(
            "<t:named-range t:name='Old' t:cell-range-address='Sheet1.A1' v:keep='yes'/>",
            "<table:named-range xmlns:table='urn:oasis:names:tc:opendocument:xmlns:table:1.0' table:name='New' table:cell-range-address='Sheet1.B2'/>",
            &["name", "cell-range-address", "base-cell-address", "range-usable-as"],
        ).unwrap();
        assert!(merged.contains("v:keep=\"yes\""));
        assert!(merged.contains("table:name=\"New\""));
    }

    #[test]
    fn bundled_named_and_database_fixtures_parse_when_available() {
        let named = std::path::Path::new(concat!(env!("CARGO_MANIFEST_DIR"), "/../../test-data/odfdo/tests/samples/simple_table_named_range.ods"));
        if named.exists() { assert!(!Spreadsheet::open(named).unwrap().named_definitions().is_empty()); }
        let database = std::path::Path::new(concat!(env!("CARGO_MANIFEST_DIR"), "/../../test-data/libreoffice-core/sc/qa/extras/testdocuments/ScDatabaseRangeObj.ods"));
        if database.exists() { assert!(!Spreadsheet::open(database).unwrap().database_ranges().is_empty()); }
    }
}
