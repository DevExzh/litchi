//! Namespace-aware ODS annotation scanning and source-span edits.

use super::{model::Cell, model::Entry, validation};
use litchi_core::{Error, Result};
use litchi_odf_common::annotation::{Annotation, Builder as AnnotationBuilder};
use quick_xml::{
    Decoder, XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::collections::BTreeMap;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const DRAW_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const SVG_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const DC_NS: &str = "http://purl.org/dc/elements/1.1/";
const META_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
const XLINK_NS: &str = "http://www.w3.org/1999/xlink";
const LOEXT_NS: &str = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
const FO_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Table,
    Other,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Kind {
    Root,
    Body,
    Spreadsheet,
    Table,
    Row,
    Cell,
    Annotation,
    Other,
}

#[derive(Clone, Debug)]
pub(crate) struct Span {
    pub(crate) start: usize,
    pub(crate) end: usize,
    pub(crate) close_start: Option<usize>,
    pub(crate) qname: String,
}

#[derive(Clone, Debug)]
pub(crate) struct Site {
    pub(crate) cell: Cell,
    pub(crate) span: Span,
    pub(crate) rows_repeated: usize,
    pub(crate) columns_repeated: usize,
}

impl Site {
    pub(crate) fn contains(&self, cell: &Cell) -> bool {
        self.cell.sheet() == cell.sheet()
            && cell.row() >= self.cell.row()
            && cell.row().saturating_sub(self.cell.row()) < self.rows_repeated
            && cell.column() >= self.cell.column()
            && cell.column().saturating_sub(self.cell.column()) < self.columns_repeated
    }
}

struct Record {
    cell: Cell,
    span: Span,
    annotation: Option<Annotation>,
}

struct Active {
    record: usize,
    builder: AnnotationBuilder,
}

struct Open {
    kind: Kind,
    span: Span,
    namespace_changes: Vec<(String, Option<String>)>,
}

struct TableState {
    name: String,
    next_row: usize,
}

struct RowState {
    start_row: usize,
    rows_repeated: usize,
    next_column: usize,
}

struct CellState {
    cell: Cell,
    rows_repeated: usize,
    columns_repeated: usize,
    span: Span,
}

/// Parsed semantic entries plus private source sites used by transactions.
pub(crate) struct Parsed {
    pub(crate) entries: Vec<Entry>,
    pub(crate) sites: Vec<Site>,
    pub(crate) annotation_spans: Vec<Span>,
}

pub(crate) fn parse(xml: &str) -> Result<Parsed> {
    validation::validate_source(xml)?;
    crate::authoring::validate_content_xml(xml)?;

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::<Open>::new();
    let mut namespaces = BTreeMap::<String, String>::new();
    let mut sheet_names = Vec::<String>::new();
    let mut current_table = None;
    let mut current_row = None;
    let mut current_cell = None;
    let mut sites = Vec::new();
    let mut records = Vec::<Record>::new();
    let mut active: Option<Active> = None;
    let mut events = 0usize;

    loop {
        let event_start = position(&reader)?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid ODS annotation XML: {error}"))
            })?;
        let namespace = namespace_kind(&resolved)?;
        let event_end = position(&reader)?;
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("ODS annotation event count overflow"))?;
        if events > validation::MAX_EVENTS {
            return Err(invalid("ODS annotation XML exceeds the event limit"));
        }
        let event = event.into_owned();

        match event {
            Event::Start(element) => {
                let changes =
                    apply_namespace_declarations(&element, reader.decoder(), &mut namespaces)?;
                let local = element.local_name();
                let local = local.as_ref();
                let span = Span {
                    start: event_start,
                    end: event_end,
                    close_start: None,
                    qname: qname(element.name().as_ref())?,
                };

                if let Some(active) = active.as_mut() {
                    if namespace == NamespaceKind::Office && local == b"annotation" {
                        return Err(invalid("nested ODS cell annotations are not supported"));
                    }
                    active.builder.start(&element, reader.decoder())?;
                    stack.push(Open {
                        kind: Kind::Other,
                        span,
                        namespace_changes: changes,
                    });
                    continue;
                }

                let parent = stack.last().map(|open| open.kind);
                let kind = classify(namespace, local, parent)?;
                match kind {
                    Kind::Table => {
                        begin_table(&element, &reader, &mut sheet_names, &mut current_table)?;
                    },
                    Kind::Row => {
                        begin_row(&element, &reader, &current_table, &mut current_row)?;
                    },
                    Kind::Cell => {
                        begin_cell(
                            &element,
                            &reader,
                            &current_table,
                            &mut current_row,
                            &mut current_cell,
                            span.clone(),
                        )?;
                    },
                    Kind::Annotation => {
                        begin_annotation(
                            &element,
                            &reader,
                            &namespaces,
                            &current_cell,
                            &mut records,
                            &mut active,
                            span.clone(),
                        )?;
                    },
                    Kind::Root | Kind::Body | Kind::Spreadsheet | Kind::Other => {},
                }
                stack.push(Open {
                    kind,
                    span,
                    namespace_changes: changes,
                });
            },
            Event::Empty(element) => {
                let changes =
                    apply_namespace_declarations(&element, reader.decoder(), &mut namespaces)?;
                let local = element.local_name();
                let local = local.as_ref();
                let span = Span {
                    start: event_start,
                    end: event_end,
                    close_start: None,
                    qname: qname(element.name().as_ref())?,
                };

                if let Some(active) = active.as_mut() {
                    if namespace == NamespaceKind::Office && local == b"annotation" {
                        return Err(invalid("nested ODS cell annotations are not supported"));
                    }
                    active.builder.empty(&element, reader.decoder())?;
                    restore_namespace_declarations(&mut namespaces, changes);
                    continue;
                }

                let parent = stack.last().map(|open| open.kind);
                let kind = classify(namespace, local, parent)?;
                match kind {
                    Kind::Table => {
                        begin_table(&element, &reader, &mut sheet_names, &mut current_table)?;
                        current_table = None;
                    },
                    Kind::Row => {
                        begin_row(&element, &reader, &current_table, &mut current_row)?;
                        finish_row(&mut current_table, &mut current_row)?;
                    },
                    Kind::Cell => {
                        begin_cell(
                            &element,
                            &reader,
                            &current_table,
                            &mut current_row,
                            &mut current_cell,
                            span.clone(),
                        )?;
                        finish_cell(&mut current_cell, &mut current_row, &mut sites)?;
                    },
                    Kind::Annotation => {
                        let cell = cell_for_annotation(&current_cell)?;
                        let builder =
                            AnnotationBuilder::new(&element, reader.decoder(), namespaces.clone())?;
                        records.push(Record {
                            cell,
                            span,
                            annotation: Some(builder.finish()?),
                        });
                    },
                    Kind::Root | Kind::Body | Kind::Spreadsheet | Kind::Other => {},
                }
                restore_namespace_declarations(&mut namespaces, changes);
            },
            Event::Text(text) => {
                if let Some(active) = active.as_mut() {
                    active.builder.text(&text)?;
                }
            },
            Event::CData(text) => {
                if let Some(active) = active.as_mut() {
                    active.builder.cdata(&text)?;
                }
            },
            Event::GeneralRef(reference) => {
                if let Some(active) = active.as_mut() {
                    active.builder.reference(&reference)?;
                }
            },
            Event::End(_) => {
                let mut open = stack
                    .pop()
                    .ok_or_else(|| invalid("ODS annotation XML element stack underflow"))?;
                open.span.end = event_end;
                if open.kind != Kind::Other {
                    open.span.close_start = Some(event_start);
                }

                if open.kind == Kind::Annotation {
                    let active = active
                        .take()
                        .ok_or_else(|| invalid("ODS annotation close has no active body"))?;
                    let record = records
                        .get_mut(active.record)
                        .ok_or_else(|| invalid("ODS annotation record is missing"))?;
                    record.span = open.span.clone();
                    record.annotation = Some(active.builder.finish()?);
                } else if let Some(active) = active.as_mut() {
                    active.builder.end_element()?;
                }

                restore_namespace_declarations(&mut namespaces, open.namespace_changes);
                match open.kind {
                    Kind::Cell => {
                        if let Some(current) = current_cell.as_mut() {
                            current.span = open.span;
                        }
                        finish_cell(&mut current_cell, &mut current_row, &mut sites)?;
                    },
                    Kind::Row => finish_row(&mut current_table, &mut current_row)?,
                    Kind::Table => {
                        if current_row.is_some() || current_cell.is_some() {
                            return Err(invalid("ODS annotation table closed before its row/cell"));
                        }
                        current_table = None;
                    },
                    Kind::Root
                    | Kind::Body
                    | Kind::Spreadsheet
                    | Kind::Annotation
                    | Kind::Other => {},
                }
            },
            Event::Eof => break,
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) | Event::DocType(_) => {},
        }
        buffer.clear();
    }

    if active.is_some()
        || !stack.is_empty()
        || current_table.is_some()
        || current_row.is_some()
        || current_cell.is_some()
    {
        return Err(invalid("ODS annotation XML ended with an unfinished owner"));
    }
    if records.len() > validation::MAX_ANNOTATIONS {
        return Err(invalid("ODS annotation count exceeds the safety limit"));
    }

    let annotation_spans = records.iter().map(|record| record.span.clone()).collect();
    let entries = records
        .into_iter()
        .enumerate()
        .map(|(index, record)| {
            let annotation = record
                .annotation
                .ok_or_else(|| invalid("ODS annotation body is unterminated"))?;
            Ok(Entry::new(index, record.cell, annotation))
        })
        .collect::<Result<Vec<_>>>()?;
    validation::validate_entries(&entries)?;
    Ok(Parsed {
        entries,
        sites,
        annotation_spans,
    })
}

fn classify(namespace: NamespaceKind, local: &[u8], parent: Option<Kind>) -> Result<Kind> {
    if namespace == NamespaceKind::Office && local == b"document-content" {
        return Ok(Kind::Root);
    }
    if namespace == NamespaceKind::Office && local == b"body" {
        return Ok(Kind::Body);
    }
    if namespace == NamespaceKind::Office && local == b"spreadsheet" {
        return Ok(Kind::Spreadsheet);
    }
    if namespace == NamespaceKind::Table && local == b"table" {
        return direct(Kind::Table, parent, Kind::Spreadsheet, "table:table");
    }
    if namespace == NamespaceKind::Table && local == b"table-row" {
        return direct(Kind::Row, parent, Kind::Table, "table:table-row");
    }
    if namespace == NamespaceKind::Table
        && (local == b"table-cell" || local == b"covered-table-cell")
    {
        return direct(Kind::Cell, parent, Kind::Row, "table cell");
    }
    if namespace == NamespaceKind::Office && local == b"annotation" {
        return direct(Kind::Annotation, parent, Kind::Cell, "office:annotation");
    }
    Ok(Kind::Other)
}

fn direct(kind: Kind, parent: Option<Kind>, expected: Kind, name: &str) -> Result<Kind> {
    if parent != Some(expected) {
        return Err(Error::InvalidFormat(format!(
            "ODS {name} is outside its worksheet context"
        )));
    }
    Ok(kind)
}

fn begin_table(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    sheet_names: &mut Vec<String>,
    current: &mut Option<TableState>,
) -> Result<()> {
    let name = required_attribute(reader, element, TABLE_NS, b"name", "table:name")?;
    if sheet_names.iter().any(|existing| existing == &name) {
        return Err(Error::InvalidFormat(format!(
            "duplicate ODS worksheet name '{name}'"
        )));
    }
    if current.is_some() {
        return Err(invalid("nested ODS worksheet tables are not supported"));
    }
    sheet_names.push(name.clone());
    *current = Some(TableState { name, next_row: 0 });
    Ok(())
}

fn begin_row(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    table: &Option<TableState>,
    current: &mut Option<RowState>,
) -> Result<()> {
    let table = table
        .as_ref()
        .ok_or_else(|| invalid("ODS row has no worksheet table"))?;
    if current.is_some() {
        return Err(invalid("nested ODS worksheet rows are not supported"));
    }
    let rows_repeated = positive_attribute(
        reader,
        element,
        b"number-rows-repeated",
        "table:number-rows-repeated",
    )?
    .unwrap_or(1);
    let start_row = table.next_row;
    start_row
        .checked_add(rows_repeated)
        .ok_or_else(|| invalid("ODS row repetition overflows coordinates"))?;
    *current = Some(RowState {
        start_row,
        rows_repeated,
        next_column: 0,
    });
    Ok(())
}

fn begin_cell(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    table: &Option<TableState>,
    row: &mut Option<RowState>,
    current: &mut Option<CellState>,
    span: Span,
) -> Result<()> {
    let table = table
        .as_ref()
        .ok_or_else(|| invalid("ODS cell has no worksheet table"))?;
    let row = row
        .as_mut()
        .ok_or_else(|| invalid("ODS cell has no worksheet row"))?;
    if current.is_some() {
        return Err(invalid("nested ODS worksheet cells are not supported"));
    }
    let columns_repeated = positive_attribute(
        reader,
        element,
        b"number-columns-repeated",
        "table:number-columns-repeated",
    )?
    .unwrap_or(1);
    let cell = Cell::new(table.name.clone(), row.start_row, row.next_column)?;
    row.next_column = row
        .next_column
        .checked_add(columns_repeated)
        .ok_or_else(|| invalid("ODS cell repetition overflows coordinates"))?;
    *current = Some(CellState {
        cell,
        rows_repeated: row.rows_repeated,
        columns_repeated,
        span,
    });
    Ok(())
}

fn finish_cell(
    current: &mut Option<CellState>,
    row: &mut Option<RowState>,
    sites: &mut Vec<Site>,
) -> Result<()> {
    let current = current
        .take()
        .ok_or_else(|| invalid("ODS worksheet cell closed without an open cell"))?;
    if row.is_none() {
        return Err(invalid("ODS worksheet cell closed without an open row"));
    }
    sites.push(Site {
        cell: current.cell,
        span: current.span,
        rows_repeated: current.rows_repeated,
        columns_repeated: current.columns_repeated,
    });
    Ok(())
}

fn finish_row(table: &mut Option<TableState>, row: &mut Option<RowState>) -> Result<()> {
    let row = row
        .take()
        .ok_or_else(|| invalid("ODS worksheet row closed without an open row"))?;
    let table = table
        .as_mut()
        .ok_or_else(|| invalid("ODS worksheet row closed without a table"))?;
    table.next_row = table
        .next_row
        .checked_add(row.rows_repeated)
        .ok_or_else(|| invalid("ODS row repetition overflows coordinates"))?;
    Ok(())
}

fn begin_annotation(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    namespaces: &BTreeMap<String, String>,
    current_cell: &Option<CellState>,
    records: &mut Vec<Record>,
    active: &mut Option<Active>,
    span: Span,
) -> Result<()> {
    let cell = cell_for_annotation(current_cell)?;
    let current = current_cell
        .as_ref()
        .ok_or_else(|| invalid("ODS annotation is not inside a worksheet cell"))?;
    if current.rows_repeated != 1 || current.columns_repeated != 1 {
        return Err(invalid(
            "ODS annotations cannot be attached to repeated logical cells",
        ));
    }
    let record = records.len();
    if record >= validation::MAX_ANNOTATIONS {
        return Err(invalid("ODS annotation count exceeds the safety limit"));
    }
    let builder = AnnotationBuilder::new(element, reader.decoder(), namespaces.clone())?;
    records.push(Record {
        cell,
        span,
        annotation: None,
    });
    *active = Some(Active { record, builder });
    Ok(())
}

fn cell_for_annotation(current: &Option<CellState>) -> Result<Cell> {
    let current = current
        .as_ref()
        .ok_or_else(|| invalid("ODS annotation is not attached to a worksheet cell"))?;
    Ok(current.cell.clone())
}

fn namespace_kind(resolved: &ResolveResult<'_>) -> Result<NamespaceKind> {
    match resolved {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_NS => Ok(NamespaceKind::Office),
        ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_NS => Ok(NamespaceKind::Table),
        ResolveResult::Bound(_) | ResolveResult::Unbound => Ok(NamespaceKind::Other),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unbound ODS annotation element prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn apply_namespace_declarations(
    element: &BytesStart<'_>,
    decoder: Decoder,
    namespaces: &mut BTreeMap<String, String>,
) -> Result<Vec<(String, Option<String>)>> {
    let mut changes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid ODS annotation namespace: {error}"))
        })?;
        let raw_name = attribute.key.as_ref();
        let prefix = if raw_name == b"xmlns" {
            String::new()
        } else if let Some(prefix) = raw_name.strip_prefix(b"xmlns:") {
            String::from_utf8(prefix.to_vec())
                .map_err(|_error| invalid("invalid XML namespace prefix"))?
        } else {
            continue;
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODS namespace URI: {error}")))?
            .into_owned();
        let previous = namespaces.insert(prefix.clone(), value);
        changes.push((prefix, previous));
    }
    Ok(changes)
}

fn restore_namespace_declarations(
    namespaces: &mut BTreeMap<String, String>,
    changes: Vec<(String, Option<String>)>,
) {
    for (prefix, previous) in changes.into_iter().rev() {
        match previous {
            Some(previous) => {
                namespaces.insert(prefix, previous);
            },
            None => {
                namespaces.remove(&prefix);
            },
        }
    }
}

fn required_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
    label: &str,
) -> Result<String> {
    attribute_value(reader, element, namespace, local)?
        .ok_or_else(|| Error::InvalidFormat(format!("ODS annotation {label} is missing")))
}

fn positive_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
    label: &str,
) -> Result<Option<usize>> {
    let Some(value) = attribute_value(reader, element, TABLE_NS, local)? else {
        return Ok(None);
    };
    let value = value
        .parse::<usize>()
        .map_err(|_error| Error::InvalidFormat(format!("ODS {label} must be positive")))?;
    if value == 0 {
        return Err(Error::InvalidFormat(format!(
            "ODS {label} must be positive"
        )));
    }
    Ok(Some(value))
}

fn attribute_value(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    expected_local: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid ODS annotation attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == *expected_namespace)
            || local.as_ref() != expected_local
        {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid ODS annotation attribute value: {error}"))
            })?
            .into_owned();
        return Ok(Some(value));
    }
    Ok(None)
}

fn qname(bytes: &[u8]) -> Result<String> {
    String::from_utf8(bytes.to_vec())
        .map_err(|_error| invalid("ODS annotation XML qualified name is not UTF-8"))
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_error| invalid("ODS annotation XML position overflows usize"))
}

pub(crate) fn serialize(annotation: &Annotation) -> Result<String> {
    validation::validate_annotation(annotation)?;
    let mut annotation = annotation.clone();
    for (prefix, uri) in [
        ("office", OFFICE_NS),
        ("text", TEXT_NS.as_bytes()),
        ("table", TABLE_NS),
        ("draw", DRAW_NS.as_bytes()),
        ("svg", SVG_NS.as_bytes()),
        ("dc", DC_NS.as_bytes()),
        ("meta", META_NS.as_bytes()),
        ("xlink", XLINK_NS.as_bytes()),
        ("loext", LOEXT_NS.as_bytes()),
        ("fo", FO_NS.as_bytes()),
        ("style", STYLE_NS.as_bytes()),
    ] {
        let uri = std::str::from_utf8(uri)
            .map_err(|_error| invalid("ODS annotation namespace URI is not UTF-8"))?;
        annotation.set_namespace(prefix, uri)?;
    }
    let mut output = String::new();
    annotation.write_xml(&mut output);
    Ok(output)
}

pub(crate) fn insert(source: &str, site: &Site, fragment: &str) -> Result<String> {
    if let Some(at) = site.span.close_start {
        return apply_edit(source, at, at, fragment);
    }
    let raw = source
        .get(site.span.start..site.span.end)
        .ok_or_else(|| invalid("invalid ODS annotation cell span"))?;
    let slash = raw
        .rfind("/>")
        .ok_or_else(|| invalid("ODS annotation cell is not self-closing or paired"))?;
    let replacement = format!("{}>{}</{}>", &raw[..slash], fragment, site.span.qname);
    apply_edit(source, site.span.start, site.span.end, &replacement)
}

pub(crate) fn remove(source: &str, span: &Span) -> Result<String> {
    apply_edit(source, span.start, span.end, "")
}

pub(crate) fn replace(source: &str, span: &Span, fragment: &str) -> Result<String> {
    apply_edit(source, span.start, span.end, fragment)
}

fn apply_edit(source: &str, start: usize, end: usize, replacement: &str) -> Result<String> {
    if start > end || end > source.len() {
        return Err(invalid("invalid ODS annotation source edit span"));
    }
    let mut output = String::with_capacity(
        source
            .len()
            .saturating_sub(end.saturating_sub(start))
            .saturating_add(replacement.len()),
    );
    output.push_str(&source[..start]);
    output.push_str(replacement);
    output.push_str(&source[end..]);
    Ok(output)
}

pub(crate) fn find_site<'a>(sites: &'a [Site], cell: &Cell) -> Option<&'a Site> {
    sites.iter().find(|site| site.contains(cell))
}

pub(crate) fn fingerprint(source: &str) -> u64 {
    // FNV-1a keeps source checks deterministic and allocation-free.
    let mut hash = 0xcbf29ce484222325u64;
    for byte in source.as_bytes() {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100000001b3);
    }
    hash
}

fn invalid(message: &str) -> Error {
    Error::InvalidFormat(message.to_string())
}
