//! Parser for data-pilot grouping declarations.

use litchi_core::Result;
use quick_xml::{events::BytesStart, events::Event, reader::NsReader};

use crate::model::data_pilot::model::{Group, GroupBoundary, GroupBy, Groups};

use super::super::xml::{is_table, optional_attr, required_attr, required_f64, text_is_whitespace};
use super::support::{invalid, parse_empty_children, xml_error};
use crate::model::data_pilot::invalid_message;

pub(super) fn parse_groups(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Groups> {
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
                groups.push(parse_group(reader, element)?);
            },
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-groups") => {
                break;
            },
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid_message("unterminated table:data-pilot-groups")),
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {
                return Err(invalid_message("invalid child in table:data-pilot-groups"));
            },
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
            .map_err(|_error| invalid("group boundary", &value)),
        (None, Some(value)) if value == "auto" => Ok(GroupBoundary::AutomaticDate),
        (None, Some(value)) => Ok(GroupBoundary::Date(value)),
    }
}
