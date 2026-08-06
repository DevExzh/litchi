//! Parser for data-pilot source declarations.

use litchi_core::Result;
use quick_xml::{
    events::{BytesStart, Event},
    reader::NsReader,
};

use crate::model::data_pilot::model::Source;
use crate::model::database_range::parse_filter;

use super::super::xml::{is_table, optional_attr, required_attr, set_once, text_is_whitespace};
use super::support::xml_error;
use crate::model::data_pilot::invalid_message;

pub(super) fn parse_cell_range_source(
    reader: &mut NsReader<&[u8]>,
    start: &BytesStart<'_>,
) -> Result<Source> {
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
            Event::End(ref element) if is_table(&namespace, element, b"source-cell-range") => {
                break;
            },
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
