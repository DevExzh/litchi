//! ODS host navigation and bounded DataPilot XML replacement.

use crate::model::data_pilot::{Table, write_data_pilot_tables};
use litchi_core::{Error, Result};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TABLE_EXT_NAMESPACE: &[u8] =
    b"urn:org:documentfoundation:names:experimental:office:xmlns:table:1.0";
const CALC_EXT_NAMESPACE: &[u8] =
    b"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";
const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_EVENTS: usize = 1_000_000;
const MAX_DEPTH: usize = 512;
const MAX_ATTRIBUTE_BYTES: usize = 64 * 1024;

/// A checked byte span for one XML element.
#[derive(Debug, Clone)]
pub(crate) struct Span {
    pub(crate) start: usize,
    pub(crate) tag_end: usize,
    pub(crate) close_start: usize,
    pub(crate) end: usize,
    pub(crate) empty: bool,
    pub(crate) qname: String,
}

/// The legal spreadsheet host and its optional DataPilot child.
#[derive(Debug, Clone)]
pub(crate) struct Location {
    pub(crate) spreadsheet: Span,
    pub(crate) container: Option<Span>,
    pub(crate) insert_at: usize,
    pub(crate) opaque: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Document,
    Body,
    Spreadsheet,
    Container,
    Other,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Table,
    TableExt,
    CalcExt,
    Other,
}

#[derive(Debug)]
struct OpenElement {
    kind: Kind,
    start: usize,
    tag_end: usize,
    qname: String,
}

/// Locate the direct spreadsheet DataPilot owner and its safe insertion point.
pub(crate) fn locate(xml: &str) -> Result<Location> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("ODS DataPilot source exceeds the size limit"));
    }

    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<OpenElement>::new();
    let mut spreadsheet = None;
    let mut container = None;
    let mut insertion = None;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut events = 0usize;
    let mut opaque = false;

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("ODS DataPilot XML event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("ODS DataPilot source exceeds the event limit"));
        }

        let event_start = position(&reader)?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid ODS DataPilot XML: {error}")))?;
        let namespace = namespace_kind(&resolved);
        let event = event.into_owned();
        let event_end = position(&reader)?;

        match event {
            Event::Start(element) => {
                if stack.is_empty() {
                    if root_seen || root_closed {
                        return Err(invalid("ODS content.xml has more than one root element"));
                    }
                    root_seen = true;
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid("ODS DataPilot source exceeds the nesting limit"));
                }
                let parent = stack.last().map(|open| open.kind);
                let kind = classify(parent, namespace, element.local_name().as_ref());
                let inside_container = stack.iter().any(|open| open.kind == Kind::Container);
                if kind == Kind::Container || inside_container {
                    if !is_known_element(namespace, element.local_name().as_ref()) {
                        opaque = true;
                    }
                    if !validate_attributes(
                        &element,
                        reader.resolver(),
                        element.local_name().as_ref(),
                    )? {
                        opaque = true;
                    }
                }
                if parent == Some(Kind::Spreadsheet)
                    && is_insertion_anchor(namespace, element.local_name().as_ref())
                    && insertion.is_none()
                {
                    insertion = Some(event_start);
                }
                let qname = element_name(&element)?;
                stack.push(OpenElement {
                    kind,
                    start: event_start,
                    tag_end: event_end,
                    qname,
                });
            },
            Event::Empty(element) => {
                if stack.is_empty() {
                    if root_seen || root_closed {
                        return Err(invalid("ODS content.xml has more than one root element"));
                    }
                    root_seen = true;
                    root_closed = true;
                }
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid("ODS DataPilot source exceeds the nesting limit"));
                }
                let parent = stack.last().map(|open| open.kind);
                let kind = classify(parent, namespace, element.local_name().as_ref());
                let inside_container = stack.iter().any(|open| open.kind == Kind::Container);
                if kind == Kind::Container || inside_container {
                    if !is_known_element(namespace, element.local_name().as_ref()) {
                        opaque = true;
                    }
                    if !validate_attributes(
                        &element,
                        reader.resolver(),
                        element.local_name().as_ref(),
                    )? {
                        opaque = true;
                    }
                }
                if parent == Some(Kind::Spreadsheet)
                    && is_insertion_anchor(namespace, element.local_name().as_ref())
                    && insertion.is_none()
                {
                    insertion = Some(event_start);
                }
                let qname = element_name(&element)?;
                record(
                    kind,
                    Span {
                        start: event_start,
                        tag_end: event_end,
                        close_start: event_start,
                        end: event_end,
                        empty: true,
                        qname,
                    },
                    &mut spreadsheet,
                    &mut container,
                )?;
            },
            Event::End(_) => {
                let open = stack.pop().ok_or_else(|| invalid("unbalanced ODS XML"))?;
                let span = Span {
                    start: open.start,
                    tag_end: open.tag_end,
                    close_start: event_start,
                    end: event_end,
                    empty: false,
                    qname: open.qname,
                };
                record(open.kind, span, &mut spreadsheet, &mut container)?;
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Text(text) if stack.iter().any(|open| open.kind == Kind::Container) => {
                let value = text
                    .xml_content(quick_xml::XmlVersion::Explicit1_0)
                    .map_err(|error| invalid(format!("invalid ODS DataPilot text: {error}")))?;
                if !value.trim().is_empty() {
                    opaque = true;
                }
            },
            Event::Comment(_) | Event::PI(_) | Event::CData(_)
                if stack.iter().any(|open| open.kind == Kind::Container) =>
            {
                opaque = true;
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("incomplete ODS content.xml document"));
    }
    let spreadsheet =
        spreadsheet.ok_or_else(|| invalid("ODS content.xml has no office:spreadsheet host"))?;
    let insert_at = insertion.unwrap_or(spreadsheet.close_start);
    Ok(Location {
        spreadsheet,
        container,
        insert_at,
        opaque,
    })
}

fn classify(parent: Option<Kind>, namespace: NamespaceKind, local: &[u8]) -> Kind {
    if parent.is_none() && is_element(namespace, NamespaceKind::Office, local, b"document-content")
    {
        Kind::Document
    } else if parent == Some(Kind::Document)
        && is_element(namespace, NamespaceKind::Office, local, b"body")
    {
        Kind::Body
    } else if parent == Some(Kind::Body)
        && is_element(namespace, NamespaceKind::Office, local, b"spreadsheet")
    {
        Kind::Spreadsheet
    } else if parent == Some(Kind::Spreadsheet)
        && is_element(namespace, NamespaceKind::Table, local, b"data-pilot-tables")
    {
        Kind::Container
    } else {
        Kind::Other
    }
}

fn record(
    kind: Kind,
    span: Span,
    spreadsheet: &mut Option<Span>,
    container: &mut Option<Span>,
) -> Result<()> {
    match kind {
        Kind::Spreadsheet => {
            if spreadsheet.replace(span).is_some() {
                return Err(invalid(
                    "ODS content.xml has more than one office:spreadsheet host",
                ));
            }
        },
        Kind::Container => {
            if container.replace(span).is_some() {
                return Err(invalid(
                    "ODS content.xml has duplicate table:data-pilot-tables",
                ));
            }
        },
        _ => {},
    }
    Ok(())
}

/// Replace only the owned DataPilot container, or insert/remove it as a
/// direct child of the spreadsheet host.
pub(crate) fn replace(
    source: &str,
    location: &Location,
    tables: Option<&[Table]>,
) -> Result<String> {
    let fragment = tables.map(render).transpose()?;
    if let Some(container) = &location.container {
        return splice(
            source,
            container.start,
            container.end,
            fragment.as_deref().unwrap_or_default(),
        );
    }
    let Some(fragment) = fragment else {
        return Ok(source.to_owned());
    };
    let spreadsheet = &location.spreadsheet;
    if spreadsheet.empty {
        let opening = source
            .get(spreadsheet.start..spreadsheet.tag_end)
            .ok_or_else(|| invalid("invalid ODS spreadsheet XML span"))?;
        let opening = opening
            .strip_suffix("/>")
            .ok_or_else(|| invalid("empty ODS spreadsheet has no close token"))?;
        let mut expanded = String::with_capacity(opening.len() + fragment.len() + 32);
        expanded.push_str(opening);
        expanded.push('>');
        expanded.push_str(&fragment);
        expanded.push_str("</");
        expanded.push_str(&spreadsheet.qname);
        expanded.push('>');
        splice(source, spreadsheet.start, spreadsheet.end, &expanded)
    } else {
        splice(source, location.insert_at, location.insert_at, &fragment)
    }
}

fn render(tables: &[Table]) -> Result<String> {
    if tables.is_empty() {
        return Ok(format!(
            "<table:data-pilot-tables xmlns:table=\"{}\"/>",
            String::from_utf8_lossy(TABLE_NAMESPACE)
        ));
    }
    let mut output = String::new();
    write_data_pilot_tables(&mut output, tables)?;
    const ROOT: &str = "<table:data-pilot-tables>";
    if !output.starts_with(ROOT) {
        return Err(invalid("DataPilot writer emitted an unexpected root"));
    }
    output.insert_str(
        ROOT.len() - 1,
        &format!(
            " xmlns:table=\"{}\"",
            String::from_utf8_lossy(TABLE_NAMESPACE)
        ),
    );
    Ok(output)
}

fn validate_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    local: &[u8],
) -> Result<bool> {
    let mut known = true;
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| invalid(format!("invalid ODS DataPilot attribute: {error}")))?;
        if attribute.value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(invalid("ODS DataPilot attribute exceeds the size limit"));
        }
        let key = attribute.key.as_ref();
        if key == b"xmlns"
            || attribute
                .key
                .prefix()
                .is_some_and(|prefix| prefix.as_ref() == b"xmlns")
        {
            continue;
        }
        let (namespace, name) = resolver.resolve_attribute(attribute.key);
        if !allowed_attribute(namespace, local, name.as_ref()) {
            known = false;
        }
    }
    Ok(known)
}

fn allowed_attribute(namespace: ResolveResult<'_>, element: &[u8], attribute: &[u8]) -> bool {
    let namespace = match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == *TABLE_NAMESPACE => NamespaceKind::Table,
        ResolveResult::Bound(Namespace(uri)) if *uri == *TABLE_EXT_NAMESPACE => {
            NamespaceKind::TableExt
        },
        ResolveResult::Bound(Namespace(uri)) if *uri == *CALC_EXT_NAMESPACE => {
            NamespaceKind::CalcExt
        },
        _ => NamespaceKind::Other,
    };
    match namespace {
        NamespaceKind::Table => match element {
            b"data-pilot-table" => matches!(
                attribute,
                b"name"
                    | b"application-data"
                    | b"grand-total"
                    | b"ignore-empty-rows"
                    | b"identify-categories"
                    | b"target-range-address"
                    | b"buttons"
                    | b"show-filter-button"
                    | b"drill-down-on-double-click"
            ),
            b"database-source-sql" => {
                matches!(
                    attribute,
                    b"database-name"
                        | b"sql-statement"
                        | b"parse-sql-statement"
                        | b"parse-sql-statements"
                )
            },
            b"database-source-table" => {
                matches!(
                    attribute,
                    b"database-name" | b"database-table-name" | b"table-name"
                )
            },
            b"database-source-query" => matches!(attribute, b"database-name" | b"query-name"),
            b"source-service" => matches!(
                attribute,
                b"name" | b"source-name" | b"object-name" | b"user-name" | b"password"
            ),
            b"source-cell-range" => matches!(attribute, b"name" | b"cell-range-address"),
            b"filter" => matches!(
                attribute,
                b"target-range-address"
                    | b"condition-source"
                    | b"condition-source-range-address"
                    | b"display-duplicates"
            ),
            b"filter-condition" => matches!(
                attribute,
                b"field-number" | b"value" | b"operator" | b"case-sensitive" | b"data-type"
            ),
            b"filter-set-item" => attribute == b"value",
            b"data-pilot-field" => matches!(
                attribute,
                b"source-field-name"
                    | b"orientation"
                    | b"selected-page"
                    | b"is-data-layout-field"
                    | b"function"
                    | b"used-hierarchy"
            ),
            b"data-pilot-level" => attribute == b"show-empty",
            b"data-pilot-subtotal" => attribute == b"function",
            b"data-pilot-member" => matches!(attribute, b"name" | b"display" | b"show-details"),
            b"data-pilot-display-info" => {
                matches!(
                    attribute,
                    b"enabled" | b"data-field" | b"member-count" | b"display-member-mode"
                )
            },
            b"data-pilot-sort-info" => {
                matches!(attribute, b"sort-mode" | b"data-field" | b"order")
            },
            b"data-pilot-layout-info" => {
                matches!(attribute, b"layout-mode" | b"add-empty-lines")
            },
            b"data-pilot-field-reference" => {
                matches!(
                    attribute,
                    b"field-name" | b"member-type" | b"member-name" | b"type"
                )
            },
            b"data-pilot-groups" => matches!(
                attribute,
                b"source-field-name"
                    | b"start"
                    | b"end"
                    | b"date-start"
                    | b"date-end"
                    | b"step"
                    | b"grouped-by"
            ),
            b"data-pilot-group" | b"data-pilot-group-member" => attribute == b"name",
            b"data-pilot-grand-total" => matches!(attribute, b"display" | b"orientation"),
            _ => false,
        },
        NamespaceKind::TableExt => {
            element == b"data-pilot-grand-total" && attribute == b"display-name"
        },
        NamespaceKind::CalcExt => {
            element == b"data-pilot-level" && attribute == b"repeat-item-labels"
        },
        _ => false,
    }
}

fn is_known_element(namespace: NamespaceKind, local: &[u8]) -> bool {
    match namespace {
        NamespaceKind::Table => matches!(
            local,
            b"data-pilot-tables"
                | b"data-pilot-table"
                | b"database-source-sql"
                | b"database-source-table"
                | b"database-source-query"
                | b"source-service"
                | b"source-cell-range"
                | b"filter"
                | b"filter-condition"
                | b"filter-and"
                | b"filter-or"
                | b"filter-set-item"
                | b"data-pilot-field"
                | b"data-pilot-level"
                | b"data-pilot-subtotals"
                | b"data-pilot-subtotal"
                | b"data-pilot-members"
                | b"data-pilot-member"
                | b"data-pilot-display-info"
                | b"data-pilot-sort-info"
                | b"data-pilot-layout-info"
                | b"data-pilot-field-reference"
                | b"data-pilot-groups"
                | b"data-pilot-group"
                | b"data-pilot-group-member"
        ),
        NamespaceKind::TableExt => local == b"data-pilot-grand-total",
        _ => false,
    }
}

fn is_insertion_anchor(namespace: NamespaceKind, local: &[u8]) -> bool {
    namespace == NamespaceKind::Table && matches!(local, b"tracked-changes" | b"shapes")
}

fn element_name(element: &BytesStart<'_>) -> Result<String> {
    String::from_utf8(element.name().as_ref().to_vec())
        .map_err(|_| invalid("ODS DataPilot element name is not UTF-8"))
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("ODS DataPilot XML position overflows usize"))
}

fn is_element(
    namespace: NamespaceKind,
    expected: NamespaceKind,
    local: &[u8],
    expected_local: &[u8],
) -> bool {
    namespace == expected && local == expected_local
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_NAMESPACE => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_NAMESPACE => NamespaceKind::Table,
        ResolveResult::Bound(Namespace(uri)) if *uri == TABLE_EXT_NAMESPACE => {
            NamespaceKind::TableExt
        },
        ResolveResult::Bound(Namespace(uri)) if *uri == CALC_EXT_NAMESPACE => {
            NamespaceKind::CalcExt
        },
        _ => NamespaceKind::Other,
    }
}

fn splice(source: &str, start: usize, end: usize, replacement: &str) -> Result<String> {
    if start > end
        || end > source.len()
        || !source.is_char_boundary(start)
        || !source.is_char_boundary(end)
    {
        return Err(invalid("invalid ODS DataPilot XML span"));
    }
    let mut output = String::with_capacity(source.len() - (end - start) + replacement.len());
    output.push_str(&source[..start]);
    output.push_str(replacement);
    output.push_str(&source[end..]);
    Ok(output)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
