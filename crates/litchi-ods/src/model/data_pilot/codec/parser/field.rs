//! Parser for data-pilot fields and field-level references.

use litchi_core::Result;
use quick_xml::{
    events::{BytesStart, Event},
    reader::NsReader,
};

use crate::model::data_pilot::model::{
    Field, FieldReference, Orientation, ReferenceMemberType, ReferenceType,
};

use super::super::xml::{
    is_table, optional_attr, optional_i64, required_attr, set_once, text_is_whitespace,
};
use super::{
    grouping::parse_groups,
    level::{level_from_start, parse_level},
    support::xml_error,
};
use crate::model::data_pilot::invalid_message;

pub(super) fn field_from_start(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Field> {
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

pub(super) fn parse_field(reader: &mut NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Field> {
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
            Event::End(ref element) if is_table(&namespace, element, b"data-pilot-field") => {
                break;
            },
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

fn parse_reference(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<FieldReference> {
    Ok(FieldReference {
        field_name: required_attr(reader, element, b"field-name")?,
        member_type: ReferenceMemberType::parse(&required_attr(reader, element, b"member-type")?)?,
        member_name: optional_attr(reader, element, b"member-name")?,
        reference_type: ReferenceType::parse(&required_attr(reader, element, b"type")?)?,
    })
}
