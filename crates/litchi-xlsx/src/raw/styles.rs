//! Bounded parser for the shared SpreadsheetML cell-format table.

use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::error::{Result, invalid};
use crate::raw::namespace::is_spreadsheetml_name;

// [MS-OE376] 2.1.728 limits the cellXfs collection to 65,430 records.
const MAX_CELL_FORMATS: u32 = 65_430;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Context {
    Styles,
    CellFormats,
    Format,
    Other,
}

/// Validated facts needed by the semantic shared-style facade.
#[derive(Debug)]
pub(crate) struct Catalog {
    cell_formats: u32,
}

impl Catalog {
    pub(crate) const fn len(&self) -> u32 {
        self.cell_formats
    }
}

pub(crate) fn parse(content: &[u8]) -> Result<Catalog> {
    let processed = litchi_ooxml_common::mce::process_ooxml(content)?;
    let mut reader = NsReader::from_reader(processed.as_ref());
    let mut stack = Vec::new();
    let mut closed_root = false;
    let mut seen_cell_formats = false;
    let mut declared_count = None;
    let mut actual_count = 0u32;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| invalid(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if stack.is_empty() => {
                if closed_root || !is_spreadsheetml_name(&namespace, element.name(), b"styleSheet")
                {
                    return Err(invalid(
                        "styles XML must have one SpreadsheetML styleSheet root",
                    ));
                }
                stack.push(Context::Styles);
            },
            Event::Empty(element) if stack.is_empty() => {
                if closed_root || !is_spreadsheetml_name(&namespace, element.name(), b"styleSheet")
                {
                    return Err(invalid(
                        "styles XML must have one SpreadsheetML styleSheet root",
                    ));
                }
                closed_root = true;
            },
            Event::Start(element) => {
                let parent = current(&stack)?;
                let context = start(
                    parent,
                    &namespace,
                    &element,
                    decoder,
                    &mut seen_cell_formats,
                    &mut declared_count,
                    &mut actual_count,
                )?;
                stack.push(context);
            },
            Event::Empty(element) => {
                let parent = current(&stack)?;
                start(
                    parent,
                    &namespace,
                    &element,
                    decoder,
                    &mut seen_cell_formats,
                    &mut declared_count,
                    &mut actual_count,
                )?;
            },
            Event::End(element) => {
                let ended = stack
                    .pop()
                    .ok_or_else(|| invalid("styles XML has a closing element outside its root"))?;
                if ended == Context::Styles {
                    if !is_spreadsheetml_name(&namespace, element.name(), b"styleSheet") {
                        return Err(invalid("styles XML has an invalid root closing element"));
                    }
                    closed_root = true;
                }
            },
            Event::Eof if !closed_root || !stack.is_empty() => {
                return Err(invalid(
                    "styles XML has a missing or unterminated SpreadsheetML styleSheet root",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !seen_cell_formats {
        return Err(invalid("styles XML is missing the cellXfs collection"));
    }
    if actual_count == 0 {
        return Err(invalid("styles XML cellXfs collection must not be empty"));
    }
    if let Some(declared) = declared_count
        && declared != actual_count
    {
        return Err(invalid(format!(
            "styles XML declares {declared} cell formats but contains {actual_count}"
        )));
    }
    Ok(Catalog {
        cell_formats: actual_count,
    })
}

#[allow(clippy::too_many_arguments)]
fn start(
    parent: Context,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    seen_cell_formats: &mut bool,
    declared_count: &mut Option<u32>,
    actual_count: &mut u32,
) -> Result<Context> {
    if parent == Context::Styles && is_spreadsheetml_name(namespace, element.name(), b"cellXfs") {
        if *seen_cell_formats {
            return Err(invalid("styles XML has duplicate cellXfs collections"));
        }
        *seen_cell_formats = true;
        *declared_count = count(element, decoder)?;
        return Ok(Context::CellFormats);
    }
    if parent == Context::CellFormats && is_spreadsheetml_name(namespace, element.name(), b"xf") {
        *actual_count = actual_count
            .checked_add(1)
            .filter(|count| *count <= MAX_CELL_FORMATS)
            .ok_or_else(|| {
                invalid(format!(
                    "styles XML contains more than {MAX_CELL_FORMATS} cell formats"
                ))
            })?;
        return Ok(Context::Format);
    }
    Ok(Context::Other)
}

fn count(element: &BytesStart<'_>, decoder: Decoder) -> Result<Option<u32>> {
    let Some(value) = unqualified_attribute_value(element, b"count", decoder)? else {
        return Ok(None);
    };
    let value = value
        .parse::<u32>()
        .map_err(|_| invalid(format!("invalid cellXfs count '{value}'")))?;
    if value > MAX_CELL_FORMATS {
        return Err(invalid(format!("cellXfs count exceeds {MAX_CELL_FORMATS}")));
    }
    Ok(Some(value))
}

fn current(stack: &[Context]) -> Result<Context> {
    stack
        .last()
        .copied()
        .ok_or_else(|| invalid("styles XML content appears outside its root"))
}

#[cfg(test)]
mod tests {
    use super::*;

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    #[test]
    fn counts_only_direct_cell_formats_after_mce_processing() {
        let xml = format!(
            r#"<styleSheet xmlns="{S}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:f="urn:future" mc:Ignorable="f"><cellXfs count="2"><xf/><xf><alignment/></xf></cellXfs><f:cellXfs><f:xf/></f:cellXfs></styleSheet>"#
        );
        assert_eq!(parse(xml.as_bytes()).expect("styles").len(), 2);
    }

    #[test]
    fn rejects_missing_duplicate_empty_and_mismatched_tables() {
        let cases = [
            format!(r#"<styleSheet xmlns="{S}"/>"#),
            format!(r#"<styleSheet xmlns="{S}"><cellXfs/><cellXfs><xf/></cellXfs></styleSheet>"#),
            format!(r#"<styleSheet xmlns="{S}"><cellXfs/></styleSheet>"#),
            format!(r#"<styleSheet xmlns="{S}"><cellXfs count="2"><xf/></cellXfs></styleSheet>"#),
        ];
        for xml in cases {
            assert!(parse(xml.as_bytes()).is_err(), "accepted {xml}");
        }
    }
}
