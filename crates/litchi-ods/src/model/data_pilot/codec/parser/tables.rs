//! Parser for the spreadsheet-level data-pilot table container.

use litchi_core::Result;
use quick_xml::{events::Event, reader::NsReader};

use crate::model::data_pilot::{
    MAX_DATA_PILOT_TABLES, invalid_message, model::Table, validation::validate_data_pilot_tables,
};

use super::super::xml::{is_office, is_table};
use super::{support::xml_error, table::parse_table};

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
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        buf.clear();
    }
    validate_data_pilot_tables(&tables)?;
    Ok(tables)
}
