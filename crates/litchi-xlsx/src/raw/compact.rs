//! Compact publication form for changed SpreadsheetML XML parts.

use litchi_core::xml::ReaderOrigin;
use quick_xml::Writer;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::reader::NsReader;

use super::namespace::is_spreadsheetml_name;
use super::worksheet::lane;
use crate::error::{Error, Result, allocation};

/// Re-emit changed XML without declaration/root or inter-element formatting.
///
/// Semantic text, entity events, and every `xml:space="preserve"` subtree are
/// retained. A leading UTF-8 byte-order mark, which the reader consumes before
/// its first event, is carried to the output unchanged: it is a byte the
/// producer wrote, not formatting. Callers must compare with the exact source
/// before invoking this function so unchanged producer XML keeps its OPC
/// source provenance.
pub(crate) fn changed(input: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    changed_observed(input, resource, &mut Unobserved).map(|(bytes, _admitted)| bytes)
}

/// A consumer of the events compaction emits.
trait Observer {
    /// Observe one event after it was emitted.
    fn event(&mut self, reader: &NsReader<&[u8]>, event: &Event<'_>);

    /// Account for an admitted `<sheetData>` body that was emitted without
    /// per-event observation.
    fn admitted(&mut self, body: &AdmittedBody<'_>);
}

/// The observer of parts that need no proof.
struct Unobserved;

impl Observer for Unobserved {
    fn event(&mut self, _reader: &NsReader<&[u8]>, _event: &Event<'_>) {}

    fn admitted(&mut self, _body: &AdmittedBody<'_>) {}
}

impl Observer for super::web::check::Probe {
    fn event(&mut self, reader: &NsReader<&[u8]>, event: &Event<'_>) {
        self.observe(reader, event);
    }

    fn admitted(&mut self, body: &AdmittedBody<'_>) {
        // Rows, cells and values are never web names, stay far below the
        // proof's depth bound and keep the root open, so the per-event proof
        // could only have been lost to an apostrophe in an attribute or to
        // value text that is not UTF-8.
        if body.summary.apostrophe || !body.value_text_is_utf8() {
            self.decline();
        }
    }
}

/// An admitted `<sheetData>` body and what the lane learned about it.
struct AdmittedBody<'a> {
    content: &'a [u8],
    start: usize,
    name: &'a [u8],
    summary: lane::Summary,
}

impl AdmittedBody<'_> {
    fn value_text_is_utf8(&self) -> bool {
        let Some(body) = self.content.get(self.start..self.summary.end) else {
            return false;
        };
        if body.is_ascii() {
            return true;
        }
        let mut valid = true;
        let walked = lane::walk(self.content, self.start, self.name, &mut |event| {
            if let lane::Event::Text {
                start,
                end,
                value: true,
            } = event
            {
                valid &= self
                    .content
                    .get(start..end)
                    .is_some_and(|text| std::str::from_utf8(text).is_ok());
            }
            Ok::<(), ()>(())
        });
        valid && matches!(walked, Ok(Some(summary)) if summary == self.summary)
    }
}

/// Keep the proof bound to the exact compacted bytes until web validation runs.
pub(crate) struct WorksheetOutput {
    bytes: Vec<u8>,
    web: super::web::check::Probe,
    admitted: bool,
}

impl WorksheetOutput {
    pub(crate) fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Whether the `<sheetData>` body was admitted by the benign lane, so
    /// every row and cell of the compacted output lies in its subset.
    pub(crate) fn body_admitted(&self) -> bool {
        self.admitted
    }

    pub(crate) fn into_bytes_and_web(self) -> Result<(Vec<u8>, crate::web::Bindings)> {
        let bindings = self.web.finish(&self.bytes)?;
        Ok((self.bytes, bindings))
    }
}

pub(crate) fn changed_worksheet(input: &[u8], resource: &'static str) -> Result<WorksheetOutput> {
    let mut web = super::web::check::Probe::default();
    let (bytes, admitted) = changed_observed(input, resource, &mut web)?;
    Ok(WorksheetOutput {
        bytes,
        web,
        admitted,
    })
}

/// Compact `input`, reporting whether a `<sheetData>` body took the lane.
fn changed_observed(
    input: &[u8],
    resource: &'static str,
    observer: &mut impl Observer,
) -> Result<(Vec<u8>, bool)> {
    let mut reader = NsReader::from_reader(input);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(input.len())
        .map_err(|source| allocation(resource, source))?;
    // The reader drops a leading byte-order mark before its first event;
    // carry it so the compacted part keeps the producer's mark.
    bytes.extend_from_slice(&input[..ReaderOrigin::of(input).skipped()]);
    let mut writer = Writer::new(bytes);
    let mut preserve = Vec::new();
    let mut admitted = false;
    #[cfg(not(test))]
    let lane_enabled = true;
    #[cfg(test)]
    let lane_enabled = lane::route::enabled();
    if let Some(entry) = compact_events(
        &mut reader,
        &mut writer,
        &mut preserve,
        observer,
        lane_enabled.then_some(input),
    )? {
        let preserved = preserve.last().copied().unwrap_or(false);
        match compact_body(input, entry, preserved, writer.get_mut(), observer)? {
            Some(resume) => {
                admitted = true;
                let spliced = lane::splice_without_body(input, entry.position, resume)?;
                let mut tail = NsReader::from_reader(spliced.as_slice());
                tail.config_mut().trim_text(false);
                tail.config_mut().check_end_names = true;
                lane::skip_to(&mut tail, &spliced, entry.position)?;
                if compact_events(&mut tail, &mut writer, &mut preserve, observer, None)?.is_some()
                {
                    return Err(invalid("compact lane resumed at a second entry"));
                }
            },
            None => {
                if compact_events(&mut reader, &mut writer, &mut preserve, observer, None)?
                    .is_some()
                {
                    return Err(invalid("compact lane resumed at a second entry"));
                }
            },
        }
    }
    Ok((writer.into_inner(), admitted))
}

/// Compact reader events until end of file.
///
/// With `lane` set to the reader's own input, stop right after the start tag
/// of the worksheet's `<sheetData>`: the first `SpreadsheetML` `sheetData`
/// child of a `SpreadsheetML` `worksheet` root, exactly the element the
/// worksheet parser and edit scanner treat as the sheet's body (a later one
/// is their duplicate refusal). The lane is decided once, at that element,
/// and only when its unprefixed children resolve to `SpreadsheetML` too, so
/// an admitted body is always that element's body.
fn compact_events(
    reader: &mut NsReader<&[u8]>,
    writer: &mut Writer<Vec<u8>>,
    preserve: &mut Vec<bool>,
    observer: &mut impl Observer,
    mut lane: Option<&[u8]>,
) -> Result<Option<lane::Entry>> {
    let mut root_seen = false;
    let mut worksheet_root = false;
    // Slice-backed events borrow the immutable input and are consumed before the next read.
    loop {
        let event_start = reader.buffer_position();
        let event = reader.read_event().map_err(xml_error)?;
        let mut candidate = None;
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
                let depth = preserve.len();
                let inherited = preserve.last().copied().unwrap_or(false);
                preserve.push(explicit.unwrap_or(inherited) || text_bearing(local.as_ref()));
                write_start(writer, element, false)?;
                if let Some(content) = lane {
                    let (namespace, _) = reader.resolver().resolve_element(element.name());
                    if depth == 0 {
                        // Only the first root can own the worksheet's body.
                        worksheet_root = !root_seen
                            && is_spreadsheetml_name(&namespace, element.name(), b"worksheet");
                        root_seen = true;
                    } else if depth == 1
                        && worksheet_root
                        && is_spreadsheetml_name(&namespace, element.name(), b"sheetData")
                    {
                        candidate = Some(
                            lane::children_are_spreadsheetml(reader.resolver())
                                .then(|| {
                                    lane::Entry::locate(
                                        content,
                                        event_start,
                                        element.name().as_ref(),
                                        reader.buffer_position(),
                                    )
                                })
                                .flatten(),
                        );
                    }
                }
            },
            Event::Empty(element) => {
                write_start(writer, element, true)?;
                if preserve.is_empty() {
                    root_seen = true;
                    worksheet_root = false;
                }
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
            Event::Eof => return Ok(None),
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => writer.write_event(event.borrow()).map_err(xml_error)?,
        }
        // Discarded formatting is absent from the bytes that web validation sees.
        observer.event(reader, &event);
        match candidate {
            Some(Some(entry)) => return Ok(Some(entry)),
            // The worksheet's body cannot take the lane; no later element may.
            Some(None) => lane = None,
            None => {},
        }
    }
}

/// Emit an admitted `<sheetData>` body without the reader.
///
/// Every admitted start, empty and close tag is already in the exact form
/// [`write_start`] and the writer produce, and value text is written
/// verbatim, so the compact body is the source body with its formatting
/// whitespace removed unless an ancestor preserves it. Returns the offset of
/// the body's close tag, or `None` to keep the reader.
fn compact_body(
    content: &[u8],
    entry: lane::Entry,
    preserved: bool,
    output: &mut Vec<u8>,
    observer: &mut impl Observer,
) -> Result<Option<usize>> {
    let name = entry.name(content)?;
    let Some(summary) = lane::recognize(content, entry.position, name) else {
        return Ok(None);
    };
    #[cfg(test)]
    lane::route::note_admitted(lane::route::Pass::Compact);
    let body = content
        .get(entry.position..summary.end)
        .ok_or_else(|| invalid("compact lane body lies outside its document"))?;
    if preserved || summary.whitespace == 0 {
        output.extend_from_slice(body);
    } else {
        let mut copied = entry.position;
        let walked = lane::walk(content, entry.position, name, &mut |event| {
            if let lane::Event::Text {
                start,
                end,
                value: false,
            } = event
            {
                output.extend_from_slice(&content[copied..start]);
                copied = end;
            }
            Ok::<(), Error>(())
        })?;
        if walked != Some(summary) {
            return Err(invalid("compact lane changed its admitted body"));
        }
        output.extend_from_slice(&content[copied..summary.end]);
    }
    observer.admitted(&AdmittedBody {
        content,
        start: entry.position,
        name,
        summary,
    });
    Ok(Some(summary.end))
}

fn invalid(message: &'static str) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(format!(
        "changed SpreadsheetML XML is malformed: {message}"
    )))
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

    /// Compact through both routes and require identical bytes, identical
    /// web-proof eligibility and identical final bindings or refusals.
    fn assert_lane_parity(input: &[u8]) -> bool {
        use crate::raw::worksheet::lane::route::{self, Pass};
        route::reset();
        let fast = changed_worksheet(input, "test worksheet XML");
        let admitted = route::admitted(Pass::Compact) > 0;
        let slow = route::without_lane(|| changed_worksheet(input, "test worksheet XML"));
        let text = String::from_utf8_lossy(input);
        match (fast, slow) {
            (Ok(fast), Ok(slow)) => {
                assert_eq!(fast.bytes(), slow.bytes(), "{text}");
                assert_eq!(fast.web.is_proven(), slow.web.is_proven(), "{text}");
                let fast = fast.into_bytes_and_web();
                let slow = slow.into_bytes_and_web();
                assert_eq!(format!("{fast:?}"), format!("{slow:?}"), "{text}");
            },
            (Err(fast), Err(slow)) => {
                assert_eq!(format!("{fast:?}"), format!("{slow:?}"), "{text}");
                assert_eq!(fast.to_string(), slow.to_string(), "{text}");
            },
            (fast, slow) => panic!(
                "routes disagree for {text}: {:?} vs {:?}",
                fast.map(|output| output.bytes().len()),
                slow.map(|output| output.bytes().len())
            ),
        }
        let unobserved_fast = changed(input, "test XML");
        let unobserved_slow = route::without_lane(|| changed(input, "test XML"));
        assert_eq!(
            format!("{unobserved_fast:?}"),
            format!("{unobserved_slow:?}"),
            "{text}"
        );
        admitted
    }

    #[test]
    fn lane_compaction_matches_the_reader_on_generated_worksheets() {
        use crate::raw::worksheet::lane::corpus::{Lcg, generated_body, worksheet};
        let mut random = Lcg(0xC0DE);
        let mut benign = 0usize;
        let mut hostile = 0usize;
        for index in 0..900 {
            let is_hostile = index % 3 == 0;
            let document = worksheet(&generated_body(&mut random, is_hostile));
            let admitted = usize::from(assert_lane_parity(document.as_bytes()));
            if is_hostile {
                hostile += admitted;
            } else {
                benign += admitted;
            }
        }
        assert_eq!(benign, 600, "every benign body is admitted");
        // Hostile bodies with duplicate attributes stay on the reader.
        assert!(
            hostile > 200 && hostile < 300,
            "hostile admissions {hostile}"
        );
    }

    #[test]
    fn lane_compaction_keeps_preserved_formatting_and_proof_inputs() {
        let cases = [
            // Formatting is dropped between rows and cells but kept in values.
            (
                format!(
                    "<worksheet xmlns=\"{MAIN_NAMESPACE}\">\n <sheetData>\n  <row r=\"1\">\n   <c r=\"A1\"> <v> 1 </v> </c>\n  </row>\n </sheetData>\n</worksheet>"
                ),
                true,
            ),
            // An inherited `xml:space` keeps the formatting.
            (
                format!(
                    "<worksheet xmlns=\"{MAIN_NAMESPACE}\" xml:space=\"preserve\"><sheetData>\n <row r=\"1\">\n </row>\n</sheetData></worksheet>"
                ),
                true,
            ),
            (
                format!(
                    "<worksheet xmlns=\"{MAIN_NAMESPACE}\"><sheetData xml:space=\"preserve\"> <row r=\"1\"/> </sheetData></worksheet>"
                ),
                true,
            ),
            // An apostrophe in an attribute value gives up the web proof on
            // both routes.
            (
                format!(
                    "<worksheet xmlns=\"{MAIN_NAMESPACE}\"><sheetData><row r=\"1\"><c r=\"A1\" x=\"it's\"/></row></sheetData></worksheet>"
                ),
                true,
            ),
            // Prefixed documents and foreign namespaces compact identically;
            // the lane takes only a SpreadsheetML worksheet body whose
            // unprefixed children resolve to SpreadsheetML.
            (
                format!(
                    "<x:worksheet xmlns:x=\"{MAIN_NAMESPACE}\"><x:sheetData><row r=\"1\"/></x:sheetData></x:worksheet>"
                ),
                false,
            ),
            (
                format!(
                    "<x:worksheet xmlns:x=\"{MAIN_NAMESPACE}\"><x:sheetData xmlns=\"{MAIN_NAMESPACE}\"><row r=\"1\"/></x:sheetData></x:worksheet>"
                ),
                true,
            ),
            (
                "<worksheet xmlns=\"urn:other\"><sheetData><row/></sheetData></worksheet>"
                    .to_owned(),
                false,
            ),
            // A web extension after the body still reaches the full reader.
            (
                format!(
                    "<worksheet xmlns=\"{MAIN_NAMESPACE}\"><sheetData><row r=\"1\"/></sheetData><extLst><ext uri=\"x\"/></extLst></worksheet>"
                ),
                true,
            ),
            // Malformed markup after the body keeps its refusal.
            (
                format!(
                    "<worksheet xmlns=\"{MAIN_NAMESPACE}\"><sheetData><row r=\"1\"/></sheetData><bad></worksheet>"
                ),
                true,
            ),
            // A nested sheetData is not the root's child.
            (
                format!(
                    "<worksheet xmlns=\"{MAIN_NAMESPACE}\"><a><sheetData><row/></sheetData></a></worksheet>"
                ),
                false,
            ),
            // Declined bodies stay on the reader.
            (
                format!(
                    "<worksheet xmlns=\"{MAIN_NAMESPACE}\"><sheetData><row r=\"1\"><c r=\"A1\"><f>1</f></c></row></sheetData></worksheet>"
                ),
                false,
            ),
        ];
        for (document, expected) in cases {
            assert_eq!(
                assert_lane_parity(document.as_bytes()),
                expected,
                "{document}"
            );
        }
    }

    #[test]
    fn lane_compaction_matches_invalid_utf8_value_text() {
        let mut document = format!(
            "<worksheet xmlns=\"{MAIN_NAMESPACE}\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"str\"><v>xx</v></c></row></sheetData></worksheet>"
        )
        .into_bytes();
        let at = document
            .windows(2)
            .position(|window| window == b"xx")
            .expect("marker");
        document[at] = 0xFF;
        assert!(assert_lane_parity(&document));
        // Non-ASCII UTF-8 text keeps the proof on both routes.
        let document = format!(
            "<worksheet xmlns=\"{MAIN_NAMESPACE}\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"str\"><v>\u{e9}</v></c></row></sheetData></worksheet>"
        );
        assert!(assert_lane_parity(document.as_bytes()));
    }

    #[test]
    fn lane_compaction_carries_the_byte_order_mark_of_marked_worksheets() {
        use crate::raw::worksheet::lane::corpus::{Lcg, generated_body, worksheet};
        let mut random = Lcg(0xB0D);
        for _ in 0..50 {
            let plain = worksheet(&generated_body(&mut random, false));
            let document = format!("\u{feff}{plain}");
            assert!(assert_lane_parity(document.as_bytes()), "{document}");
            let compacted = changed_worksheet(document.as_bytes(), "test XML").expect("compact");
            assert!(compacted.body_admitted());
            // The marked output is the unmarked output behind the mark.
            let unmarked = changed_worksheet(plain.as_bytes(), "test XML").expect("compact");
            assert_eq!(compacted.bytes().get(..3), Some(&b"\xEF\xBB\xBF"[..]));
            assert_eq!(&compacted.bytes()[3..], unmarked.bytes(), "{document}");
            assert_eq!(
                changed(document.as_bytes(), "test XML").expect("compact"),
                compacted.bytes()
            );
        }
        // A second mark is character data outside the root; compaction keeps
        // refusing or emitting it exactly as the reader reports it, after the
        // carried first mark.
        let doubled = format!("\u{feff}\u{feff}{}", worksheet("<row r=\"1\"/>"));
        let _ = assert_lane_parity(doubled.as_bytes());
    }

    #[test]
    fn lane_compaction_admits_only_the_worksheet_body() {
        let benign = "<row r=\"2\"><c r=\"A2\"><v>1</v></c></row>";
        let declined = "<row r=\"2\"><c r=\"A2\"><f>1+1</f><v>2</v></c><c r=\"B2\" t=\"inlineStr\"><is><t>x</t></is></c></row>";
        let foreign = [
            "<x:sheetData xmlns:x=\"urn:foreign\"><row r=\"1\"/></x:sheetData>",
            "<sheetData xmlns=\"urn:foreign\"><row r=\"1\"/></sheetData>",
        ];
        let admitted = |document: &str| {
            assert_lane_parity(document.as_bytes());
            changed_worksheet(document.as_bytes(), "test XML")
                .expect("compact")
                .body_admitted()
        };
        for foreign in foreign {
            // A foreign sheetData before the worksheet's own body is never
            // taken for it: the admission describes the worksheet's body.
            assert!(!admitted(&format!(
                "<worksheet xmlns=\"{MAIN_NAMESPACE}\">{foreign}<sheetData>{declined}</sheetData></worksheet>"
            )));
            assert!(admitted(&format!(
                "<worksheet xmlns=\"{MAIN_NAMESPACE}\">{foreign}<sheetData>{benign}</sheetData></worksheet>"
            )));
        }
        for document in [
            // The root is not a SpreadsheetML worksheet.
            format!(
                "<workbook xmlns=\"{MAIN_NAMESPACE}\"><sheetData>{benign}</sheetData></workbook>"
            ),
            format!("<worksheet xmlns=\"urn:foreign\"><sheetData>{benign}</sheetData></worksheet>"),
            // A nested sheetData is not the root's child.
            format!(
                "<worksheet xmlns=\"{MAIN_NAMESPACE}\"><a><sheetData>{benign}</sheetData></a></worksheet>"
            ),
            // The worksheet's body resolves its children elsewhere, and no
            // later sheetData can stand in for it.
            format!(
                "<x:worksheet xmlns:x=\"{MAIN_NAMESPACE}\"><x:sheetData xmlns=\"urn:other\">{benign}</x:sheetData><x:sheetData xmlns=\"{MAIN_NAMESPACE}\">{benign}</x:sheetData></x:worksheet>"
            ),
            // Only the first root can own the body.
            format!(
                "<worksheet xmlns=\"{MAIN_NAMESPACE}\"/><worksheet xmlns=\"{MAIN_NAMESPACE}\"><sheetData>{benign}</sheetData></worksheet>"
            ),
        ] {
            assert!(!admitted(&document), "{document}");
        }
    }
}
