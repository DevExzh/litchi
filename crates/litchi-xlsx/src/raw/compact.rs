//! Compact publication form for changed SpreadsheetML XML parts.

use quick_xml::Writer;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::reader::NsReader;

use crate::error::{Error, Result, allocation};

/// Re-emit changed XML without declaration/root or inter-element formatting.
///
/// Semantic text, entity events, and every `xml:space="preserve"` subtree are
/// retained. Callers must compare with the exact source before invoking this
/// function so unchanged producer XML keeps its OPC source provenance.
pub(crate) fn changed(input: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(input);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(input.len())
        .map_err(|source| allocation(resource, source))?;
    let mut writer = Writer::new(bytes);
    let mut preserve = Vec::new();
    loop {
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        match &event {
            Event::Start(element) => {
                let local = element.local_name();
                let mut explicit = None;
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(xml_error)?;
                    if attribute.key.as_ref() == b"xml:space" {
                        explicit = Some(attribute.value.as_ref() == b"preserve");
                    }
                }
                let inherited = preserve.last().copied().unwrap_or(false);
                preserve.push(explicit.unwrap_or(inherited) || text_bearing(local.as_ref()));
                write_start(&mut writer, element, false)?;
            },
            Event::Empty(element) => {
                write_start(&mut writer, element, true)?;
            },
            Event::End(element) => {
                let _ = preserve.pop();
                let qualified_name = element.name();
                let name = std::str::from_utf8(qualified_name.as_ref()).map_err(xml_error)?;
                writer
                    .write_event(Event::End(BytesEnd::new(name)))
                    .map_err(xml_error)?;
            },
            Event::Text(text)
                if text.as_ref().iter().all(u8::is_ascii_whitespace)
                    && !preserve.last().copied().unwrap_or(false) => {},
            Event::Eof => break,
            _ => writer.write_event(event).map_err(xml_error)?,
        }
    }
    Ok(writer.into_inner())
}

fn write_start(writer: &mut Writer<Vec<u8>>, element: &BytesStart<'_>, empty: bool) -> Result<()> {
    let qualified_name = element.name();
    let name = std::str::from_utf8(qualified_name.as_ref()).map_err(xml_error)?;
    let mut normalized = BytesStart::new(name);
    for attribute in element.attributes().with_checks(true) {
        normalized.push_attribute(attribute.map_err(xml_error)?);
    }
    writer
        .write_event(if empty {
            Event::Empty(normalized)
        } else {
            Event::Start(normalized)
        })
        .map_err(xml_error)
}

fn text_bearing(local: &[u8]) -> bool {
    matches!(
        local,
        b"t" | b"f"
            | b"v"
            | b"text"
            | b"formula"
            | b"formula1"
            | b"formula2"
            | b"sqref"
            | b"definedName"
            | b"oddHeader"
            | b"oddFooter"
            | b"evenHeader"
            | b"evenFooter"
            | b"firstHeader"
            | b"firstFooter"
    )
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(format!(
        "changed SpreadsheetML XML is malformed: {error}"
    )))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn changed_xml_is_compact_without_losing_semantic_whitespace_or_entities() {
        let input = br#"<?xml version="1.0"?>
<workbook  xmlns = "urn:test">
  <sheets/>
  <definedName>&amp;A1</definedName>
  <keep xml:space="preserve">  <future/>  </keep>
  <t>   </t>
</workbook>"#;
        let output = changed(input, "test XML").unwrap();
        let output = std::str::from_utf8(&output).unwrap();
        assert!(!output.contains("?>\n<"));
        assert!(output.contains(r#"<workbook xmlns="urn:test"><sheets/>"#));
        assert!(output.contains("<definedName>&amp;A1</definedName>"));
        assert!(output.contains(r#"<keep xml:space="preserve">  <future/>  </keep>"#));
        assert!(output.contains("<t>   </t>"));
    }
}
