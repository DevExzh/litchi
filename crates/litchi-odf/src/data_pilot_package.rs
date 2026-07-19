//! Atomic, lossless package mutation for ODS data-pilot declarations.

use crate::core::OwnedPackage;
use crate::embedded_chart::rebuild_package;
use crate::ods::data_pilot::{
    parse_data_pilot_tables, validate_data_pilot_tables, write_data_pilot_table_fragment,
};
use crate::{DataPilotGrandTotal, DataPilotTable};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{Reader, XmlVersion, events::Event, name::{Namespace, ResolveResult}, reader::NsReader};

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 128;

#[derive(Clone, Copy, PartialEq, Eq)]
enum XmlNamespace { Office, Table, Other }

/// Partial top-level metadata update. `None` retains; `Some(None)` clears.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct DataPilotTableUpdate {
    pub name: Option<String>,
    pub application_data: Option<Option<String>>,
    pub grand_total: Option<Option<DataPilotGrandTotal>>,
    pub ignore_empty_rows: Option<Option<bool>>,
    pub identify_categories: Option<Option<bool>>,
    pub target_range_address: Option<String>,
    pub buttons: Option<Option<String>>,
    pub show_filter_button: Option<Option<bool>>,
    pub drill_down_on_double_click: Option<Option<bool>>,
}

impl DataPilotTableUpdate {
    fn apply(&self, table: &mut DataPilotTable) {
        if let Some(value) = &self.name { table.name = value.clone(); }
        if let Some(value) = &self.application_data { table.application_data = value.clone(); }
        if let Some(value) = self.grand_total { table.grand_total = value; }
        if let Some(value) = self.ignore_empty_rows { table.ignore_empty_rows = value; }
        if let Some(value) = self.identify_categories { table.identify_categories = value; }
        if let Some(value) = &self.target_range_address { table.target_range_address = value.clone(); }
        if let Some(value) = &self.buttons { table.buttons = value.clone(); }
        if let Some(value) = self.show_filter_button { table.show_filter_button = value; }
        if let Some(value) = self.drill_down_on_double_click { table.drill_down_on_double_click = value; }
    }
}

#[derive(Clone)]
struct Span { start: usize, end: usize, close_start: Option<usize>, qname: String }

struct Scan {
    container: Option<Span>,
    tables: Vec<Span>,
    insertion: usize,
}

pub(crate) fn add(
    package: &OwnedPackage, content: &str, table: &DataPilotTable,
) -> Result<(Vec<u8>, usize)> {
    let (updated, index) = add_xml(content, table)?;
    rebuild(package, &updated).map(|bytes| (bytes, index))
}

pub(crate) fn replace(
    package: &OwnedPackage, content: &str, index: usize, table: &DataPilotTable,
) -> Result<Vec<u8>> {
    rebuild(package, &replace_xml(content, index, table)?)
}

pub(crate) fn update(
    package: &OwnedPackage, content: &str, index: usize, update: &DataPilotTableUpdate,
) -> Result<Vec<u8>> {
    rebuild(package, &update_xml(content, index, update)?)
}

pub(crate) fn remove(package: &OwnedPackage, content: &str, index: usize) -> Result<Vec<u8>> {
    rebuild(package, &remove_xml(content, index)?)
}

pub(crate) fn reorder(
    package: &OwnedPackage, content: &str, from: usize, to: usize,
) -> Result<Vec<u8>> {
    rebuild(package, &reorder_xml(content, from, to)?)
}

fn add_xml(content: &str, table: &DataPilotTable) -> Result<(String, usize)> {
    let current = parse_data_pilot_tables(content)?;
    let mut proposed = current.clone();
    proposed.push(table.clone());
    validate_data_pilot_tables(&proposed)?;
    let scan = scan(content)?;
    let fragment = write_data_pilot_table_fragment(table)?;
    let updated = if let Some(container) = &scan.container {
        insert_child(content, container, &fragment)?
    } else {
        let container = format!(
            "<table:data-pilot-tables xmlns:table=\"{TABLE}\">{fragment}</table:data-pilot-tables>"
        );
        splice(content, scan.insertion, scan.insertion, &container)?
    };
    validate_updated(&updated)?;
    Ok((updated, current.len()))
}

fn replace_xml(content: &str, index: usize, table: &DataPilotTable) -> Result<String> {
    let mut proposed = parse_data_pilot_tables(content)?;
    let len = proposed.len();
    let slot = proposed.get_mut(index).ok_or_else(|| bounds(index, len))?;
    *slot = table.clone();
    validate_data_pilot_tables(&proposed)?;
    let scan = scan(content)?;
    let span = scan.tables.get(index).ok_or_else(|| bounds(index, scan.tables.len()))?;
    let updated = splice(content, span.start, span.end, &write_data_pilot_table_fragment(table)?)?;
    validate_updated(&updated)?;
    Ok(updated)
}

fn update_xml(content: &str, index: usize, update: &DataPilotTableUpdate) -> Result<String> {
    let mut proposed = parse_data_pilot_tables(content)?;
    let len = proposed.len();
    update.apply(proposed.get_mut(index).ok_or_else(|| bounds(index, len))?);
    validate_data_pilot_tables(&proposed)?;
    let replacement_fragment = write_data_pilot_table_fragment(&proposed[index])?;
    let scan = scan(content)?;
    let span = scan.tables.get(index).ok_or_else(|| bounds(index, scan.tables.len()))?;
    let original_start = start_tag(&content[span.start..span.end])?;
    let start_end = span.start.checked_add(original_start.len())
        .ok_or_else(|| invalid_error("data-pilot XML position overflow"))?;
    let replacement_start = start_tag(&replacement_fragment)?;
    let merged = merge_start_tag(original_start, &span.qname, replacement_start)?;
    let updated = splice(content, span.start, start_end, &merged)?;
    validate_updated(&updated)?;
    Ok(updated)
}

fn remove_xml(content: &str, index: usize) -> Result<String> {
    let scan = scan(content)?;
    let span = scan.tables.get(index).ok_or_else(|| bounds(index, scan.tables.len()))?;
    let updated = splice(content, span.start, span.end, "")?;
    validate_updated(&updated)?;
    Ok(updated)
}

fn reorder_xml(content: &str, from: usize, to: usize) -> Result<String> {
    let scan = scan(content)?;
    let first = scan.tables.get(from).ok_or_else(|| bounds(from, scan.tables.len()))?;
    let second = scan.tables.get(to).ok_or_else(|| bounds(to, scan.tables.len()))?;
    if first.start == second.start { return Ok(content.to_string()); }
    let (left, right) = if first.start < second.start { (first, second) } else { (second, first) };
    let mut updated = String::with_capacity(content.len());
    updated.push_str(&content[..left.start]);
    updated.push_str(&content[right.start..right.end]);
    updated.push_str(&content[left.end..right.start]);
    updated.push_str(&content[left.start..left.end]);
    updated.push_str(&content[right.end..]);
    validate_updated(&updated)?;
    Ok(updated)
}

fn scan(xml: &str) -> Result<Scan> {
    if xml.len() > MAX_XML { return invalid("data-pilot content exceeds XML limit"); }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut spreadsheet_depth = None;
    let mut spreadsheet_close = None;
    let mut tail_start = None;
    let mut container_depth = None;
    let mut active_container = None;
    let mut active_table: Option<(usize, usize, String)> = None;
    let mut container = None;
    let mut tables = Vec::new();
    loop {
        let start = position(&reader)?;
        let (resolved_namespace, event) = reader.read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid_error(format!("invalid data-pilot host XML: {error}")))?;
        let namespace = classify_namespace(&resolved_namespace);
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                if is_office(&namespace, &element, b"spreadsheet") {
                    spreadsheet_depth = Some(depth);
                } else if spreadsheet_depth.is_some_and(|value| depth == value + 1)
                    && is_table(&namespace, &element, b"data-pilot-tables")
                {
                    if active_container.is_some() || container.is_some() { return invalid("duplicate data-pilot container"); }
                    container_depth = Some(depth);
                    active_container = Some((start, qname(element.name().as_ref())?));
                } else if container_depth.is_some_and(|value| depth == value + 1)
                    && is_table(&namespace, &element, b"data-pilot-table")
                {
                    if active_table.is_some() { return invalid("nested data-pilot table"); }
                    active_table = Some((depth, start, qname(element.name().as_ref())?));
                } else if spreadsheet_depth.is_some_and(|value| depth == value + 1)
                    && tail_start.is_none()
                    && (is_table(&namespace, &element, b"consolidation") || is_table(&namespace, &element, b"dde-links"))
                {
                    tail_start = Some(start);
                }
                depth = depth.checked_add(1).ok_or_else(|| invalid_error("data-pilot XML depth overflow"))?;
                if depth > MAX_DEPTH { return invalid("data-pilot XML nesting exceeds limit"); }
            },
            Event::Empty(element) => {
                if spreadsheet_depth.is_some_and(|value| depth == value + 1)
                    && is_table(&namespace, &element, b"data-pilot-tables")
                {
                    if active_container.is_some() || container.is_some() { return invalid("duplicate data-pilot container"); }
                    container = Some(Span { start, end, close_start: None, qname: qname(element.name().as_ref())? });
                } else if container_depth.is_some_and(|value| depth == value + 1)
                    && is_table(&namespace, &element, b"data-pilot-table")
                {
                    tables.push(Span { start, end, close_start: None, qname: qname(element.name().as_ref())? });
                } else if spreadsheet_depth.is_some_and(|value| depth == value + 1)
                    && tail_start.is_none()
                    && (is_table(&namespace, &element, b"consolidation") || is_table(&namespace, &element, b"dde-links"))
                {
                    tail_start = Some(start);
                }
            },
            Event::End(element) => {
                depth = depth.checked_sub(1).ok_or_else(|| invalid_error("data-pilot XML depth underflow"))?;
                if active_table.as_ref().is_some_and(|(at, _, _)| *at == depth)
                    && is_table_end(&namespace, &element, b"data-pilot-table")
                {
                    let (_, table_start, name) = active_table.take().expect("active data-pilot table");
                    tables.push(Span { start: table_start, end, close_start: Some(start), qname: name });
                } else if container_depth == Some(depth) && is_table_end(&namespace, &element, b"data-pilot-tables") {
                    let (container_start, name) = active_container.take().ok_or_else(|| invalid_error("missing data-pilot container start"))?;
                    container = Some(Span { start: container_start, end, close_start: Some(start), qname: name });
                    container_depth = None;
                } else if spreadsheet_depth == Some(depth) && is_office_end(&namespace, &element, b"spreadsheet") {
                    spreadsheet_close = Some(start);
                    spreadsheet_depth = None;
                }
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in data-pilot host XML"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if active_container.is_some() || active_table.is_some() { return invalid("unterminated data-pilot structure"); }
    tables.sort_by_key(|span| span.start);
    let insertion = tail_start.or(spreadsheet_close).ok_or_else(|| invalid_error("office:spreadsheet host was not found"))?;
    Ok(Scan { container, tables, insertion })
}

fn validate_updated(xml: &str) -> Result<()> {
    scan(xml)?;
    validate_data_pilot_tables(&parse_data_pilot_tables(xml)?)
}

fn insert_child(xml: &str, container: &Span, fragment: &str) -> Result<String> {
    if let Some(close) = container.close_start {
        splice(xml, close, close, fragment)
    } else {
        let raw = &xml[container.start..container.end];
        let slash = raw.rfind("/>").ok_or_else(|| invalid_error("invalid empty data-pilot container"))?;
        let replacement = format!("{}>{}</{}>", &raw[..slash], fragment, container.qname);
        splice(xml, container.start, container.end, &replacement)
    }
}

fn start_tag(fragment: &str) -> Result<&str> {
    let end = fragment.find('>').ok_or_else(|| invalid_error("invalid data-pilot fragment"))?;
    Ok(&fragment[..=end])
}

fn merge_start_tag(original: &str, original_qname: &str, replacement: &str) -> Result<String> {
    const KNOWN: &[&str] = &[
        "name", "application-data", "grand-total", "ignore-empty-rows", "identify-categories",
        "target-range-address", "buttons", "show-filter-button", "drill-down-on-double-click",
    ];
    let original_prefix = original_qname.split_once(':').map_or("", |(prefix, _)| prefix);
    let original_attributes = read_attributes(original)?;
    let replacement_attributes = read_attributes(replacement)?;
    let mut output = format!("<{original_qname}");
    let mut written = std::collections::HashSet::new();
    for (name, value) in replacement_attributes {
        write_attr(&mut output, &name, &value);
        written.insert(name);
    }
    for (name, value) in original_attributes {
        let (prefix, local) = name.split_once(':').unwrap_or(("", name.as_str()));
        let known = (prefix == original_prefix || prefix == "table") && KNOWN.contains(&local);
        if !known && !written.contains(&name) {
            write_attr(&mut output, &name, &value);
            written.insert(name);
        }
    }
    output.push('>');
    Ok(output)
}

fn read_attributes(start_tag: &str) -> Result<Vec<(String, String)>> {
    let mut reader = Reader::from_str(start_tag);
    let mut buffer = Vec::new();
    match reader.read_event_into(&mut buffer).map_err(|error| invalid_error(format!("invalid data-pilot start tag: {error}")))? {
        Event::Start(element) | Event::Empty(element) => element.attributes().map(|attribute| {
            let attribute = attribute.map_err(|error| invalid_error(format!("invalid data-pilot attribute: {error}")))?;
            let name = qname(attribute.key.as_ref())?;
            let value = attribute.decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| invalid_error(format!("invalid data-pilot attribute value: {error}")))?.into_owned();
            Ok((name, value))
        }).collect(),
        _ => invalid("invalid data-pilot start tag"),
    }
}

fn write_attr(output: &mut String, name: &str, value: &str) {
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    output.push_str(&escape_xml(value));
    output.push('"');
}

fn splice(xml: &str, start: usize, end: usize, replacement: &str) -> Result<String> {
    if start > end || end > xml.len() || !xml.is_char_boundary(start) || !xml.is_char_boundary(end) {
        return invalid("invalid data-pilot XML splice");
    }
    let mut output = String::with_capacity(xml.len() - (end - start) + replacement.len());
    output.push_str(&xml[..start]);
    output.push_str(replacement);
    output.push_str(&xml[end..]);
    Ok(output)
}

fn rebuild(package: &OwnedPackage, content: &str) -> Result<Vec<u8>> {
    rebuild_package(package, content, Vec::new(), Vec::new(), Vec::new(), Vec::new())
}

fn classify_namespace(namespace: &ResolveResult<'_>) -> XmlNamespace {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_NS => XmlNamespace::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_NS => XmlNamespace::Table,
        _ => XmlNamespace::Other,
    }
}
fn is_table(namespace: &XmlNamespace, element: &quick_xml::events::BytesStart<'_>, local: &[u8]) -> bool {
    *namespace == XmlNamespace::Table
        && element.local_name().as_ref() == local
}
fn is_table_end(namespace: &XmlNamespace, element: &quick_xml::events::BytesEnd<'_>, local: &[u8]) -> bool {
    *namespace == XmlNamespace::Table
        && element.local_name().as_ref() == local
}
fn is_office(namespace: &XmlNamespace, element: &quick_xml::events::BytesStart<'_>, local: &[u8]) -> bool {
    *namespace == XmlNamespace::Office
        && element.local_name().as_ref() == local
}
fn is_office_end(namespace: &XmlNamespace, element: &quick_xml::events::BytesEnd<'_>, local: &[u8]) -> bool {
    *namespace == XmlNamespace::Office
        && element.local_name().as_ref() == local
}
fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_| invalid_error("data-pilot XML position overflow"))
}
fn qname(value: &[u8]) -> Result<String> {
    std::str::from_utf8(value).map(str::to_string).map_err(|_| invalid_error("invalid data-pilot qualified name"))
}
fn bounds(index: usize, len: usize) -> Error {
    invalid_error(format!("data-pilot table index {index} is out of bounds for {len} entries"))
}
fn invalid_error(message: impl Into<String>) -> Error { Error::InvalidFormat(message.into()) }
fn invalid<T>(message: impl Into<String>) -> Result<T> { Err(invalid_error(message)) }

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{DataPilotField, DataPilotOrientation, DataPilotSource};

    const PREFIX: &str = "<office:document-content xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0' xmlns:table='urn:oasis:names:tc:opendocument:xmlns:table:1.0' xmlns:v='urn:vendor'><office:body><office:spreadsheet><table:table table:name='Source'/><table:table table:name='Result'/>";
    const SUFFIX: &str = "<table:consolidation table:function='sum' table:source-cell-range-addresses='Source.A1:A2' table:target-cell-address='Result.Z1'/></office:spreadsheet></office:body></office:document-content>";

    fn pivot(name: &str, target: &str) -> DataPilotTable {
        let mut table = DataPilotTable::new(name, target);
        table.source = Some(DataPilotSource::CellRange {
            name: None,
            cell_range_address: "Source.A1:Source.C20".to_string(),
            filter: None,
        });
        table.fields.push(DataPilotField::new("Region", DataPilotOrientation::Row));
        table
    }

    #[test]
    fn generated_crud_orders_container_and_preserves_unknown_update_xml() {
        let original = format!("{PREFIX}<table:data-pilot-tables><table:data-pilot-table table:name='Old' table:target-range-address='Result.A1:C5' v:flag='keep'><v:cache/><table:data-pilot-field table:source-field-name='Region' table:orientation='row'/></table:data-pilot-table></table:data-pilot-tables>{SUFFIX}");
        let updated = update_xml(&original, 0, &DataPilotTableUpdate { name: Some("Renamed".to_string()), ..Default::default() }).unwrap();
        assert!(updated.contains("v:flag=\"keep\"") && updated.contains("<v:cache/>") && updated.contains("table:name=\"Renamed\""));
        let (added, index) = add_xml(&updated, &pivot("Second", "Result.E1:G5")).unwrap();
        assert_eq!(index, 1);
        assert!(added.find("data-pilot-tables").unwrap() < added.find("table:consolidation").unwrap());
        let reordered = reorder_xml(&added, 0, 1).unwrap();
        assert!(reordered.find("Second").unwrap() < reordered.find("Renamed").unwrap());
        let removed = remove_xml(&reordered, 1).unwrap();
        assert_eq!(parse_data_pilot_tables(&removed).unwrap().len(), 1);
    }

    #[test]
    fn rejects_duplicate_names_overlapping_targets_and_bad_references() {
        let first = pivot("P", "Result.A1:C5");
        let mut duplicate = pivot("P", "Result.E1:G5");
        assert!(validate_data_pilot_tables(&[first.clone(), duplicate.clone()]).is_err());
        duplicate.name = "Q".to_string();
        duplicate.target_range_address = "Result.C5:F8".to_string();
        assert!(validate_data_pilot_tables(&[first, duplicate]).is_err());
    }

    #[test]
    fn replace_round_trips_grouping_and_named_source() {
        let mut table = pivot("Grouped", "Result.A1:D9");
        if let Some(DataPilotSource::CellRange { name, .. }) = &mut table.source { *name = Some("SalesSource".to_string()); }
        let xml = format!("{PREFIX}<table:named-expressions><table:named-range table:name='SalesSource' table:base-cell-address='Source.A1' table:cell-range-address='Source.A1:C20'/></table:named-expressions><table:data-pilot-tables>{}</table:data-pilot-tables>{SUFFIX}", write_data_pilot_table_fragment(&pivot("Old", "Result.J1:L4")).unwrap());
        let replaced = replace_xml(&xml, 0, &table).unwrap();
        let parsed = parse_data_pilot_tables(&replaced).unwrap();
        assert!(matches!(&parsed[0].source, Some(DataPilotSource::CellRange { name: Some(value), .. }) if value == "SalesSource"));
    }
}
