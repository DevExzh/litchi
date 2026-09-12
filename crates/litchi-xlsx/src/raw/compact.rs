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
    changed_observed(input, resource, |_, _| {})
}

/// Keep the proof bound to the exact compacted bytes until web validation runs.
pub(crate) struct WorksheetOutput {
    bytes: Vec<u8>,
    web: super::web::check::Probe,
}

impl WorksheetOutput {
    pub(crate) fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    pub(crate) fn into_bytes_and_web(self) -> Result<(Vec<u8>, crate::web::Bindings)> {
        let bindings = self.web.finish(&self.bytes)?;
        Ok((self.bytes, bindings))
    }
}

pub(crate) fn changed_worksheet(input: &[u8], resource: &'static str) -> Result<WorksheetOutput> {
    let mut web = super::web::check::Probe::default();
    let bytes = changed_observed(input, resource, |reader, event| web.observe(reader, event))?;
    Ok(WorksheetOutput { bytes, web })
}

fn changed_observed(
    input: &[u8],
    resource: &'static str,
    mut observe: impl FnMut(&NsReader<&[u8]>, &Event<'_>),
) -> Result<Vec<u8>> {
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
                    && !preserve.last().copied().unwrap_or(false) =>
            {
                continue;
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => writer.write_event(event.borrow()).map_err(xml_error)?,
        }
        // Discarded formatting is absent from the bytes that web validation sees.
        observe(&reader, &event);
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

    const MAIN_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    #[derive(Debug, PartialEq, Eq)]
    enum PipelineOutcome {
        Error {
            phase: &'static str,
            bytes: Vec<u8>,
            debug: String,
            display: String,
        },
        Success {
            bytes: Vec<u8>,
            bindings: crate::web::Bindings,
        },
    }

    fn pipeline_error(phase: &'static str, bytes: Vec<u8>, error: Error) -> PipelineOutcome {
        PipelineOutcome::Error {
            phase,
            bytes,
            debug: format!("{error:?}"),
            display: error.to_string(),
        }
    }

    fn reference_pipeline(input: &[u8], parse_grid: bool) -> PipelineOutcome {
        let bytes = match changed(input, "test XML") {
            Ok(bytes) => bytes,
            Err(error) => return pipeline_error("compact", Vec::new(), error),
        };
        if parse_grid {
            if let Err(error) = crate::raw::worksheet::parse(&bytes, || Ok(None)) {
                return pipeline_error("grid", bytes, error);
            }
        }
        match crate::raw::web::read(&bytes) {
            Ok(bindings) => PipelineOutcome::Success { bytes, bindings },
            Err(error) => pipeline_error("web", bytes, error),
        }
    }

    fn fused_pipeline(input: &[u8], parse_grid: bool) -> PipelineOutcome {
        let compacted = match changed_worksheet(input, "test XML") {
            Ok(compacted) => compacted,
            Err(error) => return pipeline_error("compact", Vec::new(), error),
        };
        let bytes = compacted.bytes().to_vec();
        if parse_grid {
            if let Err(error) = crate::raw::worksheet::parse(compacted.bytes(), || Ok(None)) {
                return pipeline_error("grid", bytes, error);
            }
        }
        match compacted.into_bytes_and_web() {
            Ok((bytes, bindings)) => PipelineOutcome::Success { bytes, bindings },
            Err(error) => pipeline_error("web", bytes, error),
        }
    }

    fn assert_pipeline_parity(input: &[u8], parse_grid: bool) -> PipelineOutcome {
        let expected = reference_pipeline(input, parse_grid);
        let actual = fused_pipeline(input, parse_grid);
        assert_eq!(actual, expected);
        actual
    }

    fn ordinary_worksheet() -> Vec<u8> {
        format!(
            r#"<worksheet xmlns="{MAIN_NAMESPACE}">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>"#
        )
        .into_bytes()
    }

    #[test]
    fn changed_worksheet_fuses_the_successful_ordinary_pipeline() {
        let input = ordinary_worksheet();
        let output = assert_pipeline_parity(&input, true);
        let PipelineOutcome::Success { bytes, bindings } = output else {
            panic!("ordinary worksheet should pass both validation phases");
        };
        assert!(bindings.is_empty());
        assert!(std::str::from_utf8(&bytes)
            .expect("compacted worksheet is UTF-8")
            .contains(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#));

        let compacted = changed_worksheet(&input, "test XML").expect("compact worksheet");
        assert!(compacted.web.is_proven());
    }

    #[test]
    fn changed_worksheet_observes_only_emitted_whitespace() {
        let input = format!(
            r#"<worksheet xmlns="{MAIN_NAMESPACE}">
  <sheetData>
    <row r="1">
      <c r="A1"><v> 1 </v></c>
    </row>
  </sheetData>
</worksheet>"#
        )
        .into_bytes();
        let output = assert_pipeline_parity(&input, true);
        let PipelineOutcome::Success { bytes, bindings } = output else {
            panic!("whitespace-only worksheet should pass both validation phases");
        };
        assert!(bindings.is_empty());
        let text = std::str::from_utf8(&bytes).expect("compacted worksheet is UTF-8");
        assert!(text.contains("<v> 1 </v>"));
        assert!(!text.contains("</worksheet>\n"));
        let compacted = changed_worksheet(&input, "test XML").expect("compact worksheet");
        assert!(compacted.web.is_proven());
    }

    #[test]
    fn changed_worksheet_keeps_compact_grid_and_web_error_order() {
        let compact_error = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></worksheet>"#;
        let grid_error = format!(
            r#"<worksheet xmlns="{MAIN_NAMESPACE}"><sheetData><row r="not-a-row"/></sheetData></worksheet>"#
        );
        let web_error = format!(
            r#"<worksheet xmlns="{MAIN_NAMESPACE}" xmlns:x15="{x15}">
  <sheetData/>
  <extLst><ext uri="{uri}"><x15:webExtensions/></ext></extLst>
</worksheet>"#,
            x15 = crate::raw::web::X15_NAMESPACE,
            uri = crate::raw::web::EXTENSION_URI,
        );

        let compact_result = assert_pipeline_parity(compact_error, true);
        assert!(matches!(
            compact_result,
            PipelineOutcome::Error {
                phase: "compact",
                ..
            }
        ));

        let grid_result = assert_pipeline_parity(grid_error.as_bytes(), true);
        assert!(matches!(
            grid_result,
            PipelineOutcome::Error { phase: "grid", .. }
        ));

        let web_result = assert_pipeline_parity(web_error.as_bytes(), true);
        assert!(matches!(
            web_result,
            PipelineOutcome::Error { phase: "web", .. }
        ));

        let competing_grid_and_web_error = format!(
            r#"<worksheet xmlns="{MAIN_NAMESPACE}" xmlns:x15="{x15}">
  <sheetData><row r="not-a-row"/></sheetData>
  <extLst><ext uri="{uri}"><x15:webExtensions/></ext></extLst>
</worksheet>"#,
            x15 = crate::raw::web::X15_NAMESPACE,
            uri = crate::raw::web::EXTENSION_URI,
        );
        let competing_result =
            assert_pipeline_parity(competing_grid_and_web_error.as_bytes(), true);
        assert!(matches!(
            competing_result,
            PipelineOutcome::Error { phase: "grid", .. }
        ));

        let competing_compact_and_grid_and_web_error = format!(
            r#"<worksheet xmlns="{MAIN_NAMESPACE}" xmlns:x15="{x15}">
  <sheetData><row r="not-a-row"/></sheetData>
  <extLst><ext uri="{uri}"><x15:webExtensions/></ext></extLst>
  <tail></wrong>
</worksheet>"#,
            x15 = crate::raw::web::X15_NAMESPACE,
            uri = crate::raw::web::EXTENSION_URI,
        );
        let competing_result =
            assert_pipeline_parity(competing_compact_and_grid_and_web_error.as_bytes(), true);
        assert!(matches!(
            competing_result,
            PipelineOutcome::Error {
                phase: "compact",
                ..
            }
        ));

        let web_only_result = assert_pipeline_parity(web_error.as_bytes(), false);
        assert!(matches!(
            web_only_result,
            PipelineOutcome::Error { phase: "web", .. }
        ));
    }

    #[test]
    fn changed_worksheet_falls_back_for_extension_bindings() {
        let input = format!(
            r#"<worksheet xmlns="{MAIN_NAMESPACE}" xmlns:x15="{x15}" xmlns:xm="{xm}">
  <sheetData/>
  <extLst><ext uri="{uri}"><x15:webExtensions>
    <x15:webExtension appRef="dashboard"><xm:f>Sheet1!A1:B2</xm:f></x15:webExtension>
  </x15:webExtensions></ext></extLst>
</worksheet>"#,
            x15 = crate::raw::web::X15_NAMESPACE,
            xm = crate::raw::web::XM_NAMESPACE,
            uri = crate::raw::web::EXTENSION_URI,
        );
        let compacted = changed_worksheet(input.as_bytes(), "test XML")
            .expect("extension worksheet compaction");
        assert!(!compacted.web.is_proven());

        let output = assert_pipeline_parity(input.as_bytes(), false);
        let PipelineOutcome::Success { bindings, .. } = output else {
            panic!("valid extension binding should pass web validation");
        };
        assert_eq!(bindings.len(), 1);
        assert_eq!(
            bindings.get("dashboard").expect("binding").formula(),
            "Sheet1!A1:B2"
        );
    }

    #[test]
    fn changed_worksheet_does_not_prove_single_quoted_attributes() {
        let input = format!(
            r#"<worksheet xmlns="{MAIN_NAMESPACE}" marker='contains "quotes"'><sheetData/></worksheet>"#
        )
        .into_bytes();
        let compacted = changed_worksheet(&input, "test XML").expect("single-quoted XML scan");
        assert!(!compacted.web.is_proven());
        let _ = assert_pipeline_parity(&input, false);
    }

    #[test]
    fn changed_worksheet_can_compact_an_oversized_whitespace_input() {
        let mut input = format!(r#"<worksheet xmlns="{MAIN_NAMESPACE}">"#).into_bytes();
        input.extend(std::iter::repeat_n(b' ', 16 * 1024 * 1024 + 1));
        input.extend_from_slice(b"<sheetData/></worksheet>");
        assert!(input.len() > 16 * 1024 * 1024);

        let output = assert_pipeline_parity(&input, false);
        let PipelineOutcome::Success { bytes, bindings } = output else {
            panic!("discarded formatting should leave a bounded worksheet");
        };
        assert!(bytes.len() <= 16 * 1024 * 1024);
        assert!(bindings.is_empty());
        let compacted = changed_worksheet(&input, "test XML").expect("compact worksheet");
        assert!(compacted.web.is_proven());
    }
}
