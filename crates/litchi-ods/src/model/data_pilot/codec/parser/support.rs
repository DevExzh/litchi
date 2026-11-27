//! Shared XML structure and error handling for data-pilot parsing.

use litchi_core::Result;
use litchi_odf_common::{ResolvedReader, XmlStreamLimits};
use quick_xml::events::Event;

use super::super::xml::{is_table, text_is_whitespace};
use crate::model::data_pilot::invalid_message;

pub(super) fn invalid(kind: &str, value: &str) -> litchi_core::Error {
    invalid_message(&format!("invalid {kind} '{value}'"))
}

pub(super) fn parse_empty_children(
    reader: &mut ResolvedReader<'_>,
    parent: &[u8],
    child: &[u8],
    attribute: &[u8],
    limits: &XmlStreamLimits,
) -> Result<Vec<String>> {
    let mut values = Vec::new();
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element)
                if is_table(&namespace, element, child) =>
            {
                if values.len() >= limits.max_events() {
                    return Err(invalid_message(
                        "data-pilot child collection exceeds the configured event limit",
                    ));
                }
                values.push(super::super::xml::required_attr(
                    reader, element, attribute,
                )?);
            },
            Event::End(ref element) if is_table(&namespace, element, parent) => break,
            Event::End(ref element) if is_table(&namespace, element, child) => {},
            Event::Text(ref text) if text_is_whitespace(text)? => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid_message("unterminated data-pilot child container")),
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {
                return Err(invalid_message("invalid data-pilot child element"));
            },
        }
    }
    Ok(values)
}

pub(super) fn xml_error(error: litchi_odf_common::Error) -> litchi_core::Error {
    litchi_odf_common::map_resolved_reader_error(error, "invalid data-pilot XML")
}
