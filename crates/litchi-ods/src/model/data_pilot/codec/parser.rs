//! ODF data-pilot XML parser.

use crate::model::database_range::{
    parse_filter, parse_source_query, parse_source_sql, parse_source_table,
};
use litchi_core::Result;
use quick_xml::{
    events::{BytesStart, Event},
    reader::NsReader,
};

use super::super::{
    MAX_DATA_PILOT_TABLES, invalid_message,
    model::{
        DisplayInfo, DisplayMemberMode, Field, FieldReference, GrandTotal, GrandTotalElement,
        GrandTotalOrientation, Group, GroupBoundary, GroupBy, Groups, LayoutInfo, LayoutMode,
        Level, Member, Orientation, ReferenceMemberType, ReferenceType, SortInfo, SortMode,
        SortOrder, Source, Table,
    },
    validation::validate_data_pilot_tables,
};
use super::xml::{
    CALC_EXT_NAMESPACE, TABLE_EXT_NAMESPACE, consume_empty_extension, is_office, is_table,
    is_table_ext, is_table_namespace, optional_attr, optional_bool, optional_i64, optional_ns_attr,
    optional_ns_bool, required_attr, required_bool, required_f64, required_u64, set_once,
    skip_foreign_element, text_is_whitespace,
};

pub(crate) fn parse_data_pilot_tables(xml: &str) -> Result<Vec<Table>> {
    let mut reader = NsReader::from_str(xml);
    let mut buf = Vec::new();
    let mut depth = 0usize;
    let mut spreadsheet_depth = None;
    let mut container_depth = None;
    let mut container_seen = false;
    let mut tables = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) => {
                depth += 1;
                if is_office(&namespace, element, b"spreadsheet") {
                    spreadsheet_depth = Some(depth);
                } else if is_table(&namespace, element, b"data-pilot-tables")
                    && spreadsheet_depth.is_some_and(|value| depth == value + 1)
                {
                    if container_seen {
                        return Err(invalid_message("duplicate table:data-pilot-tables"));
                    }
                    container_seen = true;
                    container_depth = Some(depth);
                } else if is_table(&namespace, element, b"data-pilot-table")
                    && container_depth.is_some_and(|value| depth == value + 1)
                {
                    let table = parse_table(&mut reader, element)?;
                    table.validate()?;
                    tables.push(table);
                    if tables.len() > MAX_DATA_PILOT_TABLES {
                        return Err(invalid_message(
                            "data-pilot table count exceeds safety limit",
                        ));
                    }
                    depth -= 1;
                } else if container_depth.is_some() {
                    return Err(invalid_message("invalid child in table:data-pilot-tables"));
                }
            },
            Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-table")
                    && container_depth == Some(depth) =>
            {
                return Err(invalid_message(
                    "data-pilot table requires at least one field",
                ));
            },
            Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-tables")
                    && spreadsheet_depth == Some(depth) =>
            {
                if container_seen {
                    return Err(invalid_message("duplicate table:data-pilot-tables"));
                }
                container_seen = true;
            },
            Event::End(ref element) => {
                if is_table(&namespace, element, b"data-pilot-tables")
                    && container_depth == Some(depth)
                {
                    container_depth = None;
                } else if is_office(&namespace, element, b"spreadsheet")
                    && spreadsheet_depth == Some(depth)
                {
                    spreadsheet_depth = None;
                }
                depth = depth.saturating_sub(1);
            },
            Event::Eof => break,
            _ => {},
        }
        buf.clear();
    }
    validate_data_pilot_tables(&tables)?;
    Ok(tables)
}

fn parse_table(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Table> {
    let mut table = Table {
        name: required_attr(reader, start, b"name")?,
        application_data: optional_attr(reader, start, b"application-data")?,
        grand_total: optional_attr(reader, start, b"grand-total")?
            .map(|value| GrandTotal::parse(&value))
            .transpose()?,
        ignore_empty_rows: optional_bool(reader, start, b"ignore-empty-rows")?,
        identify_categories: optional_bool(reader, start, b"identify-categories")?,
        target_range_address: required_attr(reader, start, b"target-range-address")?,
        buttons: optional_attr(reader, start, b"buttons")?,
        show_filter_button: optional_bool(reader, start, b"show-filter-button")?,
        drill_down_on_double_click: optional_bool(reader, start, b"drill-down-on-double-click")?,
        grand_totals: Vec::new(),
        source: None,
        fields: Vec::new(),
    };
    let mut fields_started = false;
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element)
                if is_table_ext(&namespace, element, b"data-pilot-grand-total") =>
            {
                if fields_started || table.source.is_some() {
                    return Err(invalid_message(
                        "data-pilot grand totals must precede the source and fields",
                    ));
                }
                table.grand_totals.push(parse_grand_total(reader, element)?);
                consume_empty_extension(reader, b"data-pilot-grand-total")?;
            },
            Event::Empty(ref element)
                if is_table_ext(&namespace, element, b"data-pilot-grand-total") =>
            {
                if fields_started || table.source.is_some() {
                    return Err(invalid_message(
                        "data-pilot grand totals must precede the source and fields",
                    ));
                }
                table.grand_totals.push(parse_grand_total(reader, element)?);
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"database-source-sql") =>
            {
                if fields_started {
                    return Err(invalid_message("data-pilot source must precede all fields"));
                }
                set_once(
                    &mut table.source,
                    Source::Database(parse_source_sql(reader, element)?),
                    "data-pilot source",
                )?;
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"database-source-table") =>
            {
                if fields_started {
                    return Err(invalid_message("data-pilot source must precede all fields"));
                }
                set_once(
                    &mut table.source,
                    Source::Database(parse_source_table(reader, element)?),
                    "data-pilot source",
                )?;
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"database-source-query") =>
            {
                if fields_started {
                    return Err(invalid_message("data-pilot source must precede all fields"));
                }
                set_once(
                    &mut table.source,
                    Source::Database(parse_source_query(reader, element)?),
                    "data-pilot source",
                )?;
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"source-service") =>
            {
                if fields_started {
                    return Err(invalid_message("data-pilot source must precede all fields"));
                }
                let source = Source::Service {
                    name: required_attr(reader, element, b"name")?,
                    source_name: required_attr(reader, element, b"source-name")?,
                    object_name: required_attr(reader, element, b"object-name")?,
                    user_name: optional_attr(reader, element, b"user-name")?,
                    password: optional_attr(reader, element, b"password")?,
                };
                set_once(&mut table.source, source, "data-pilot source")?;
            },
            Event::Start(ref element) if is_table(&namespace, element, b"source-cell-range") => {
                if fields_started {
                    return Err(invalid_message("data-pilot source must precede all fields"));
                }
                let source = parse_cell_range_source(reader, element)?;
                set_once(&mut table.source, source, "data-pilot source")?;
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"source-cell-range") => {
                if fields_started {
                    return Err(invalid_message("data-pilot source must precede all fields"));
                }
                let source = Source::CellRange {
                    name: optional_attr(reader, element, b"name")?,
                    cell_range_address: required_attr(reader, element, b"cell-range-address")?,
                    filter: None,
                };
                set_once(&mut table.source, source, "data-pilot source")?;
            },
            Event::Start(ref element) if is_table(&namespace, element, b"data-pilot-field") => {
                fields_started = true;
                table.fields.push(parse_field(reader, element)?);
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"data-pilot-field") => {
                fields_started = true;
                table.fields.push(field_from_start(reader, element)?);
            },
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-table") => break,
            Event::End(ref element)
                if is_table(&namespace, element, b"database-source-sql")
                    || is_table(&namespace, element, b"database-source-table")
                    || is_table(&namespace, element, b"database-source-query")
                    || is_table(&namespace, element, b"source-service") => {},
            Event::Text(ref text) => {
                if !text_is_whitespace(text)? {
                    return Err(invalid_message("data-pilot table cannot contain text"));
                }
            },
            Event::Eof => return Err(invalid_message("unterminated table:data-pilot-table")),
            Event::Comment(_) => {},
            Event::Start(ref element) if !is_table_namespace(&namespace) => {
                skip_foreign_element(reader, element)?;
            },
            Event::Empty(_) if !is_table_namespace(&namespace) => {},
            other => {
                return Err(invalid_message(&format!(
                    "invalid child in table:data-pilot-table: {other:?}"
                )));
            },
        }
        buf.clear();
    }
    Ok(table)
}

fn parse_cell_range_source(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Source> {
    let address = required_attr(reader, start, b"cell-range-address")?;
    let name = optional_attr(reader, start, b"name")?;
    let mut filter = None;
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) if is_table(&namespace, element, b"filter") => {
                set_once(&mut filter, parse_filter(reader, element)?, "source filter")?;
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"filter") => {
                return Err(invalid_message("table:filter has no expression"));
            },
            Event::End(ref element) if is_table(&namespace, element, b"source-cell-range") => break,
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid_message("unterminated table:source-cell-range")),
            _ => return Err(invalid_message("invalid child in table:source-cell-range")),
        }
        buf.clear();
    }
    Ok(Source::CellRange {
        name,
        cell_range_address: address,
        filter,
    })
}

fn field_from_start(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Field> {
    let orientation = Orientation::parse(&required_attr(reader, start, b"orientation")?)?;
    Ok(Field {
        source_field_name: required_attr(reader, start, b"source-field-name")?,
        orientation,
        selected_page: optional_attr(reader, start, b"selected-page")?,
        is_data_layout_field: optional_attr(reader, start, b"is-data-layout-field")?,
        function: optional_attr(reader, start, b"function")?,
        used_hierarchy: optional_i64(reader, start, b"used-hierarchy")?,
        level: None,
        reference: None,
        groups: None,
    })
}

fn parse_field(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Field> {
    let mut field = field_from_start(reader, start)?;
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) if is_table(&namespace, element, b"data-pilot-level") => {
                set_once(
                    &mut field.level,
                    parse_level(reader, element)?,
                    "data-pilot level",
                )?;
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"data-pilot-level") => {
                set_once(
                    &mut field.level,
                    level_from_start(reader, element)?,
                    "data-pilot level",
                )?;
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-field-reference") =>
            {
                set_once(
                    &mut field.reference,
                    parse_reference(reader, element)?,
                    "field reference",
                )?;
            },
            Event::Start(ref element) if is_table(&namespace, element, b"data-pilot-groups") => {
                set_once(
                    &mut field.groups,
                    parse_groups(reader, element)?,
                    "field groups",
                )?;
            },
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-field") => break,
            Event::End(ref element)
                if is_table(&namespace, element, b"data-pilot-field-reference") => {},
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid_message("unterminated table:data-pilot-field")),
            _ => return Err(invalid_message("invalid child in table:data-pilot-field")),
        }
        buf.clear();
    }
    field.validate()?;
    Ok(field)
}

fn level_from_start(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Level> {
    Ok(Level {
        show_empty: optional_bool(reader, start, b"show-empty")?,
        repeat_item_labels: optional_ns_bool(
            reader,
            start,
            CALC_EXT_NAMESPACE,
            b"repeat-item-labels",
        )?,
        ..Default::default()
    })
}

fn parse_level(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Level> {
    let mut level = level_from_start(reader, start)?;
    let mut subtotals_seen = false;
    let mut members_seen = false;
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) if is_table(&namespace, element, b"data-pilot-subtotals") => {
                if std::mem::replace(&mut subtotals_seen, true) {
                    return Err(invalid_message("duplicate data-pilot subtotals"));
                }
                level.subtotals = parse_subtotals(reader)?
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"data-pilot-subtotals") => {
                if std::mem::replace(&mut subtotals_seen, true) {
                    return Err(invalid_message("duplicate data-pilot subtotals"));
                }
            },
            Event::Start(ref element) if is_table(&namespace, element, b"data-pilot-members") => {
                if std::mem::replace(&mut members_seen, true) {
                    return Err(invalid_message("duplicate data-pilot members"));
                }
                level.members = parse_members(reader)?
            },
            Event::Empty(ref element) if is_table(&namespace, element, b"data-pilot-members") => {
                if std::mem::replace(&mut members_seen, true) {
                    return Err(invalid_message("duplicate data-pilot members"));
                }
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-display-info") =>
            {
                set_once(
                    &mut level.display,
                    parse_display(reader, element)?,
                    "display info",
                )?;
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-sort-info") =>
            {
                set_once(&mut level.sort, parse_sort(reader, element)?, "sort info")?;
            },
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-layout-info") =>
            {
                set_once(
                    &mut level.layout,
                    parse_layout(reader, element)?,
                    "layout info",
                )?;
            },
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-level") => break,
            Event::End(ref element)
                if is_table(&namespace, element, b"data-pilot-display-info")
                    || is_table(&namespace, element, b"data-pilot-sort-info")
                    || is_table(&namespace, element, b"data-pilot-layout-info") => {},
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Start(ref element) if !is_table_namespace(&namespace) => {
                skip_foreign_element(reader, element)?
            },
            Event::Empty(_) if !is_table_namespace(&namespace) => {},
            Event::Eof => return Err(invalid_message("unterminated table:data-pilot-level")),
            _ => return Err(invalid_message("invalid child in table:data-pilot-level")),
        }
        buf.clear();
    }
    Ok(level)
}

fn parse_subtotals(reader: &mut NsReader<&[u8]>) -> Result<Vec<String>> {
    parse_empty_children(
        reader,
        b"data-pilot-subtotals",
        b"data-pilot-subtotal",
        b"function",
    )
}

fn parse_members(reader: &mut NsReader<&[u8]>) -> Result<Vec<Member>> {
    let mut members = Vec::new();
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, b"data-pilot-member") =>
            {
                members.push(Member {
                    name: required_attr(reader, element, b"name")?,
                    display: optional_bool(reader, element, b"display")?,
                    show_details: optional_bool(reader, element, b"show-details")?,
                })
            },
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-members") => {
                break;
            },
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-member") => {},
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid_message("unterminated table:data-pilot-members")),
            _ => return Err(invalid_message("invalid child in table:data-pilot-members")),
        }
        buf.clear();
    }
    Ok(members)
}

fn parse_display(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<DisplayInfo> {
    Ok(DisplayInfo {
        enabled: required_bool(reader, element, b"enabled")?,
        data_field: required_attr(reader, element, b"data-field")?,
        member_count: required_u64(reader, element, b"member-count")?,
        mode: DisplayMemberMode::parse(&required_attr(reader, element, b"display-member-mode")?)?,
    })
}

fn parse_sort(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<SortInfo> {
    Ok(SortInfo {
        mode: SortMode::parse(&required_attr(reader, element, b"sort-mode")?)?,
        data_field: optional_attr(reader, element, b"data-field")?,
        order: SortOrder::parse(&required_attr(reader, element, b"order")?)?,
    })
}

fn parse_layout(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<LayoutInfo> {
    Ok(LayoutInfo {
        mode: LayoutMode::parse(&required_attr(reader, element, b"layout-mode")?)?,
        add_empty_lines: required_bool(reader, element, b"add-empty-lines")?,
    })
}

fn parse_reference(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<FieldReference> {
    Ok(FieldReference {
        field_name: required_attr(reader, element, b"field-name")?,
        member_type: ReferenceMemberType::parse(&required_attr(reader, element, b"member-type")?)?,
        member_name: optional_attr(reader, element, b"member-name")?,
        reference_type: ReferenceType::parse(&required_attr(reader, element, b"type")?)?,
    })
}

fn parse_groups(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Groups> {
    let start_boundary = parse_boundary(reader, start, b"start", b"date-start")?;
    let end_boundary = parse_boundary(reader, start, b"end", b"date-end")?;
    let mut groups = Vec::new();
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) if is_table(&namespace, element, b"data-pilot-group") => {
                groups.push(parse_group(reader, element)?)
            },
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-groups") => break,
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid_message("unterminated table:data-pilot-groups")),
            _ => return Err(invalid_message("invalid child in table:data-pilot-groups")),
        }
        buf.clear();
    }
    Ok(Groups {
        source_field_name: required_attr(reader, start, b"source-field-name")?,
        start: start_boundary,
        end: end_boundary,
        step: required_f64(reader, start, b"step")?,
        grouped_by: GroupBy::parse(&required_attr(reader, start, b"grouped-by")?)?,
        groups,
    })
}

fn parse_group(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Group> {
    Ok(Group {
        name: required_attr(reader, start, b"name")?,
        members: parse_empty_children(
            reader,
            b"data-pilot-group",
            b"data-pilot-group-member",
            b"name",
        )?,
    })
}

fn parse_empty_children(
    reader: &mut NsReader<&[u8]>,
    parent: &[u8],
    child: &[u8],
    attribute: &[u8],
) -> Result<Vec<String>> {
    let mut values = Vec::new();
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, child) =>
            {
                values.push(required_attr(reader, element, attribute)?)
            },
            Event::End(ref element) if is_table(&namespace, element, parent) => break,
            Event::End(ref element) if is_table(&namespace, element, child) => {},
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid_message("unterminated data-pilot child container")),
            _ => return Err(invalid_message("invalid data-pilot child element")),
        }
        buf.clear();
    }
    Ok(values)
}

fn parse_boundary(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    numeric: &[u8],
    date: &[u8],
) -> Result<GroupBoundary> {
    match (
        optional_attr(reader, element, numeric)?,
        optional_attr(reader, element, date)?,
    ) {
        (Some(_), Some(_)) | (None, None) => Err(invalid_message(
            "data-pilot grouping requires exactly one boundary attribute",
        )),
        (Some(value), None) if value == "auto" => Ok(GroupBoundary::AutomaticNumber),
        (Some(value), None) => value
            .parse::<f64>()
            .map(GroupBoundary::Number)
            .map_err(|_| invalid("group boundary", &value)),
        (None, Some(value)) if value == "auto" => Ok(GroupBoundary::AutomaticDate),
        (None, Some(value)) => Ok(GroupBoundary::Date(value)),
    }
}

fn parse_grand_total(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<GrandTotalElement> {
    Ok(GrandTotalElement {
        orientation: GrandTotalOrientation::parse(&required_attr(
            reader,
            element,
            b"orientation",
        )?)?,
        display: required_bool(reader, element, b"display")?,
        display_name: optional_ns_attr(reader, element, TABLE_EXT_NAMESPACE, b"display-name")?,
    })
}

fn invalid(kind: &str, value: &str) -> litchi_core::Error {
    invalid_message(&format!("invalid {kind} '{value}'"))
}

fn xml_error(error: quick_xml::Error) -> litchi_core::Error {
    invalid_message(&format!("XML parsing error: {error}"))
}
