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
    // Slice-backed events borrow the immutable input and are consumed before the next read.
    loop {
        let event = reader.read_event().map_err(xml_error)?;
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
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => writer.write_event(event).map_err(xml_error)?,
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

    /// Keep the pre-experiment ownership path in the differential harness.
    /// This intentionally mirrors `changed` before it began borrowing events.
    fn changed_owned(input: &[u8], resource: &'static str) -> Result<Vec<u8>> {
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
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => writer.write_event(event).map_err(xml_error)?,
            }
        }
        Ok(writer.into_inner())
    }

    fn assert_owned_event_parity(input: &[u8]) {
        let borrowed = changed(input, "test XML");
        let owned = changed_owned(input, "test XML");
        match (borrowed, owned) {
            (Ok(actual), Ok(expected)) => assert_eq!(actual, expected),
            (Err(actual), Err(expected)) => {
                assert_eq!(format!("{actual:?}"), format!("{expected:?}"));
                assert_eq!(actual.to_string(), expected.to_string());
            },
            (actual, expected) => {
                panic!("borrowed and owned compaction paths disagree: {actual:?} vs {expected:?}")
            },
        }
    }

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

    #[test]
    fn changed_borrowed_events_match_owned_events_for_mixed_markup() {
        let inputs: &[&[u8]] = &[
            br#"<?xml version="1.0" encoding="UTF-8"?>
<x:root xmlns:x="urn:root" xmlns:q="urn:qualified" q:flag="yes" plain="&quot;v&quot;">
  <q:item q:attr="one" plain="two"/>
  <x:child xml:space="preserve">  keep &amp; entities  </x:child>
  <x:reset xml:space="default">  compact  </x:reset>
</x:root>"#,
            br#"<!DOCTYPE root><?before instruction?><root>
  <!-- preserved comment -->
  <![CDATA[raw <markup> & bytes]]>
  <t xml:space="preserve">  text &amp; &#x41; &custom;  </t>
  <?inside processing?>
</root>"#,
            br#"<root xml:space="preserve">
  <outer>
    <inner/>
  </outer>
  <outer xml:space="default">
    <inner/>
  </outer>
  <empty xmlns:q="urn:q" q:kind="empty" />
</root>"#,
        ];

        for input in inputs {
            let actual = changed(input, "test XML").expect("mixed markup should be accepted");
            let expected =
                changed_owned(input, "test XML").expect("mixed markup should be accepted");
            assert_eq!(actual, expected);
        }

        let output = changed(inputs[2], "test XML").expect("xml:space fixture should be accepted");
        assert!(
            output
                .windows(b"<outer xml:space=\"default\"><inner/></outer>".len())
                .any(|window| window == b"<outer xml:space=\"default\"><inner/></outer>")
        );
    }

    #[test]
    fn changed_borrowed_events_match_owned_events_for_all_text_bearing_forms() {
        let input = br#"<root xmlns:x="urn:x">
  <t>  </t>
  <f>
  </f>
  <v>
   </v>
  <x:text>
  </x:text>
  <formula>
 </formula>
  <x:formula>
   </x:formula>
  <formula1>
    </formula1>
  <formula2>
  </formula2>
  <sqref>
    </sqref>
  <x:definedName>
 </x:definedName>
  <oddHeader>
  </oddHeader>
  <oddFooter>
    </oddFooter>
  <evenHeader>
  </evenHeader>
  <evenFooter>
    </evenFooter>
  <firstHeader>
 </firstHeader>
  <firstFooter>
  </firstFooter>
</root>"#;

        assert_owned_event_parity(input);
        let output = changed(input, "test XML").expect("text-bearing forms should be accepted");
        for fragment in [
            "<t>  </t>",
            "<f>\n  </f>",
            "<v>\n   </v>",
            "<x:text>\n  </x:text>",
            "<formula>\n </formula>",
            "<x:formula>\n   </x:formula>",
            "<formula1>\n    </formula1>",
            "<formula2>\n  </formula2>",
            "<sqref>\n    </sqref>",
            "<x:definedName>\n </x:definedName>",
            "<oddHeader>\n  </oddHeader>",
            "<oddFooter>\n    </oddFooter>",
            "<evenHeader>\n  </evenHeader>",
            "<evenFooter>\n    </evenFooter>",
            "<firstHeader>\n </firstHeader>",
            "<firstFooter>\n  </firstFooter>",
        ] {
            assert!(
                output
                    .windows(fragment.len())
                    .any(|window| window == fragment.as_bytes()),
                "text-bearing whitespace was not retained: {fragment:?}"
            );
        }
    }

    #[test]
    fn changed_borrowed_events_match_owned_events_for_malformed_inputs() {
        let invalid_utf8_start = [b'<', 0xff, b'/', b'>'];
        let invalid_utf8_text = [b'<', b'r', b'>', 0xff, b'<', b'/', b'r', b'>'];
        let invalid_utf8_attribute = [b'<', b'r', b' ', b'a', b'=', b'"', 0xff, b'"', b'/', b'>'];
        let invalid_utf8_end = [b'<', b'r', b'>', b'<', b'/', 0xff, b'>'];
        let inputs: &[&[u8]] = &[
            br#"<root a="one" a="two"/>"#,
            br#"<root xmlns:q="urn:q" q:a="one" q:a="two"/>"#,
            br#"<root a=one/>"#,
            br#"<root a/>"#,
            br#"<root a=>"#,
            br#"<root a="unterminated></root>"#,
            br#"<root><child></root>"#,
            br#"<root/><tail>"#,
            br#"<root/><tail></wrong>"#,
            br#"<root>dangling &reference</root>"#,
            br#"<root><!-- unclosed</root>"#,
            &invalid_utf8_start,
            &invalid_utf8_text,
            &invalid_utf8_attribute,
            &invalid_utf8_end,
        ];

        for input in inputs {
            assert_owned_event_parity(input);
        }
    }
}
