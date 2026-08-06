//! Shared XML structure and error handling for data-pilot parsing.

use litchi_core::Result;
use quick_xml::{events::Event, reader::NsReader};

use super::super::xml::{is_table, text_is_whitespace};
use crate::model::data_pilot::invalid_message;

pub(super) fn invalid(kind: &str, value: &str) -> litchi_core::Error {
    invalid_message(&format!("invalid {kind} '{value}'"))
}

pub(super) fn parse_empty_children(
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
                values.push(super::super::xml::required_attr(
                    reader, element, attribute,
                )?)
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

pub(super) fn xml_error(error: quick_xml::Error) -> litchi_core::Error {
    invalid_message(&format!("XML parsing error: {error}"))
}
