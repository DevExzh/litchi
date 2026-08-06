//! Parser for one `table:data-pilot-table` declaration.

use litchi_core::Result;
use quick_xml::{
    events::{BytesStart, Event},
    reader::NsReader,
};

use crate::model::data_pilot::{
    invalid_message,
    model::{GrandTotal, GrandTotalElement, GrandTotalOrientation, Source, Table},
};
use crate::model::database_range::{parse_source_query, parse_source_sql, parse_source_table};

use super::super::xml::{
    TABLE_EXT_NAMESPACE, consume_empty_extension, is_table, is_table_ext, is_table_namespace,
    optional_attr, optional_bool, optional_ns_attr, required_attr, required_bool, set_once,
    skip_foreign_element, text_is_whitespace,
};
use super::{
    field::{field_from_start, parse_field},
    source::parse_cell_range_source,
    support::xml_error,
};

pub(super) fn parse_table(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Table> {
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
