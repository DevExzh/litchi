//! Focused regression tests for the worksheet codec seams.

use std::mem::size_of;

use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

use super::wire::{cell_tag, tag as owned_tag, write_cell_tag};
use super::{Attribute, Tag, scan, scan_with_event_limit, sibling_name, write_tag};
use crate::raw::worksheet::model::MAX_XML_DEPTH;
use crate::raw::worksheet::parse_a1;

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

#[test]
fn sibling_names_retain_the_source_prefix() {
    assert_eq!(sibling_name("x:row", "c"), "x:c");
    assert_eq!(sibling_name("row", "c"), "c");
}

#[test]
fn wire_tags_escape_and_replace_attributes_deterministically() {
    let tag = Tag {
        name: "x:c".into(),
        attributes: vec![Attribute {
            name: "r".into(),
            value: "A1".into(),
        }]
        .into_boxed_slice(),
    };
    let mut output = Vec::new();
    write_tag(&mut output, &tag, true, &["r"], &[("r", "A&B".to_owned())]);
    assert_eq!(output, br#"<x:c r="A&amp;B"/>"#);
}

#[test]
fn snapshot_scan_builds_sorted_edit_slots() {
    let source = format!(
        r#"<x:worksheet xmlns:x="{MAIN}"><x:dimension ref="A1:B2"/><x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row><x:row r="2"><x:c r="B2"/></x:row></x:sheetData></x:worksheet>"#
    );
    let layout = scan(source.as_bytes()).expect("worksheet scan");
    assert_eq!(layout.sheet_data.rows.len(), 2);
    assert_eq!(layout.sheet_data.rows[0].cells.len(), 1);
    assert_eq!(layout.sheet_data.rows[1].cells[0].address.a1(), "B2");
}

#[test]
fn snapshot_scan_restores_namespace_scope_after_a_rebinding() {
    let source = format!(
        r#"<x:worksheet xmlns:x="{MAIN}"><x:sheetData><x:row r="1"><x:c r="A1"/></x:row><x:row xmlns:x="urn:future" r="2"><x:c r="B2"/></x:row><x:row r="3"><x:c r="C3"/></x:row></x:sheetData></x:worksheet>"#
    );
    let layout = scan(source.as_bytes()).expect("worksheet scan");
    assert_eq!(
        layout
            .sheet_data
            .rows
            .iter()
            .map(|row| row.number)
            .collect::<Vec<_>>(),
        [1, 3]
    );
}

#[test]
fn snapshot_scan_rejects_nesting_beyond_worksheet_depth_limit() {
    let mut source = format!(r#"<worksheet xmlns="{MAIN}">"#);
    for _ in 0..MAX_XML_DEPTH {
        source.push_str("<future>");
    }
    for _ in 0..MAX_XML_DEPTH {
        source.push_str("</future>");
    }
    source.push_str("</worksheet>");

    let error = scan(source.as_bytes()).expect_err("deep worksheet should be rejected");
    assert_eq!(
        error.to_string(),
        format!("invalid XLSX structure: worksheet XML exceeds {MAX_XML_DEPTH} levels")
    );
}

#[test]
fn snapshot_scan_accepts_nesting_at_worksheet_depth_limit() {
    let mut source = format!(r#"<worksheet xmlns="{MAIN}"><sheetData/>"#);
    for _ in 0..(MAX_XML_DEPTH - 1) {
        source.push_str("<future>");
    }
    for _ in 0..(MAX_XML_DEPTH - 1) {
        source.push_str("</future>");
    }
    source.push_str("</worksheet>");

    let layout = scan(source.as_bytes()).expect("boundary-depth worksheet should be accepted");
    assert!(layout.sheet_data.empty);
}

#[test]
fn snapshot_scan_rejects_mismatched_end_name() {
    let source = format!(r#"<worksheet xmlns="{MAIN}"><sheetData></wrong></worksheet>"#);
    assert!(scan(source.as_bytes()).is_err());
}

#[test]
fn snapshot_scan_handles_large_flat_event_stream_within_depth_limit() {
    let mut source = format!(r#"<worksheet xmlns="{MAIN}"><sheetData/>"#);
    for _ in 0..(MAX_XML_DEPTH * 16) {
        source.push_str("<future/>");
    }
    source.push_str("</worksheet>");

    let layout = scan(source.as_bytes()).expect("flat worksheet events should be scanned");
    assert!(layout.sheet_data.empty);
}

#[test]
fn snapshot_scan_accepts_event_count_at_limit() {
    let source = format!(r#"<worksheet xmlns="{MAIN}"><sheetData/></worksheet>"#);
    let layout = scan_with_event_limit(source.as_bytes(), 4)
        .expect("worksheet event count at the limit should be accepted");
    assert!(layout.sheet_data.empty);
}

#[test]
fn snapshot_scan_rejects_flat_event_stream_over_event_limit() {
    let mut source = format!(r#"<worksheet xmlns="{MAIN}"><sheetData/>"#);
    for _ in 0..64 {
        source.push_str("<future/>");
    }
    source.push_str("</worksheet>");

    let error = scan_with_event_limit(source.as_bytes(), 8)
        .expect_err("flat worksheet event stream should be bounded");
    assert_eq!(
        error.to_string(),
        "invalid XLSX structure: worksheet XML exceeds event limit"
    );
}

fn start_element(xml: &[u8]) -> (BytesStart<'static>, Decoder) {
    let mut reader = NsReader::from_reader(xml);
    let event = reader.read_event().expect("cell start event");
    let element = match event {
        Event::Start(element) | Event::Empty(element) => element.into_owned(),
        other => panic!("expected a cell start event, got {other:?}"),
    };
    (element, reader.decoder())
}

fn serialize_cell_tag(xml: &[u8], expected_reference: &str) -> (Option<Tag>, Vec<u8>, Vec<u8>) {
    let (old_element, old_decoder) = start_element(xml);
    let old = owned_tag(&old_element, old_decoder).expect("owned cell tag");
    let mut old_output = Vec::new();
    write_tag(
        &mut old_output,
        &old,
        true,
        &["r"],
        &[("r", expected_reference.to_owned())],
    );

    let (new_element, new_decoder) = start_element(xml);
    let compact = cell_tag(&new_element, new_decoder).expect("compact cell tag");
    let mut new_output = Vec::new();
    match compact.as_ref() {
        Some(tag) => write_tag(
            &mut new_output,
            tag,
            true,
            &["r"],
            &[("r", expected_reference.to_owned())],
        ),
        None => write_cell_tag(
            &mut new_output,
            true,
            &[("r", expected_reference.to_owned())],
        ),
    }
    (compact, old_output, new_output)
}

#[test]
fn compact_cell_tag_uses_the_niche_without_growing_the_owned_tag() {
    assert!(size_of::<Option<Tag>>() <= size_of::<Tag>());
}

#[test]
fn compact_cell_tag_matches_owned_serialization_for_plain_cells() {
    let cases = [
        (br#"<c/>"#.as_slice(), false),
        (br#"<c r="A1"/>"#.as_slice(), false),
        (br#"<c r="A&#x31;"/>"#.as_slice(), false),
        (br#"<c r=" A1 "/>"#.as_slice(), false),
        (br#"<c r="A1" s="7"/>"#.as_slice(), true),
        (br#"<c s="7" r="A1"/>"#.as_slice(), true),
        (br#"<c cm="1" vm="2" t="inlineStr"/>"#.as_slice(), true),
        (br#"<c r="A1" future="a&amp;b"/>"#.as_slice(), true),
        (br#"<x:c r="A1"/>"#.as_slice(), true),
        (
            br#"<c xmlns:x="urn:future" x:opaque="yes" r="A1"/>"#.as_slice(),
            true,
        ),
    ];

    for (xml, retains_tag) in cases {
        let (compact, old_output, new_output) = serialize_cell_tag(xml, "A1");
        assert_eq!(
            compact.is_some(),
            retains_tag,
            "unexpected tag retention for {xml:?}"
        );
        assert_eq!(new_output, old_output, "serialization changed for {xml:?}");
    }
}

#[test]
fn compact_cell_writer_matches_empty_style_and_clear_start_shapes() {
    let cases = [
        (br#"<c/>"#.as_slice(), true, vec![("r", "A1")]),
        (
            br#"<c r="A1"/>"#.as_slice(),
            true,
            vec![("r", "A1"), ("s", "7")],
        ),
        (
            br#"<c r="A1"/>"#.as_slice(),
            true,
            vec![("r", "A1"), ("t", "inlineStr")],
        ),
        (br#"<c r="A1">"#.as_slice(), false, vec![("r", "A1")]),
        (
            br#"<c r="A1">"#.as_slice(),
            false,
            vec![("r", "A1"), ("t", "inlineStr"), ("future", "A&B")],
        ),
    ];
    for (xml, empty, attributes) in cases {
        let (element, decoder) = start_element(xml);
        let old = owned_tag(&element, decoder).expect("owned cell tag");
        let appended = attributes
            .into_iter()
            .map(|(name, value)| (name, value.to_owned()))
            .collect::<Vec<_>>();
        let mut expected = Vec::new();
        write_tag(
            &mut expected,
            &old,
            empty,
            &["r", "s", "t", "future"],
            &appended,
        );
        let mut output = Vec::new();
        write_cell_tag(&mut output, empty, &appended);
        assert_eq!(output, expected, "compact writer changed {xml:?}");
    }
}

#[test]
fn snapshot_scan_marks_only_plain_cells_as_tagless() {
    let source = format!(
        r#"<worksheet xmlns="{MAIN}" xmlns:x="{MAIN}" xmlns:future="urn:future"><sheetData><row r="1"><c/><c r="B1"/><c r="C1" s="7" cm="1" vm="2" future:opaque="yes"/><x:c r="D1"/><c r="E1" xmlns:future="urn:future" future:opaque="yes"/></row></sheetData></worksheet>"#
    );
    let layout = scan(source.as_bytes()).expect("worksheet scan");
    let cells = &layout.sheet_data.rows[0].cells;
    assert_eq!(cells.len(), 5);
    assert!(cells[0].tag.is_none());
    assert!(cells[1].tag.is_none());
    assert!(cells[2].tag.is_some());
    assert!(cells[3].tag.is_some());
    assert!(cells[4].tag.is_some());
}

#[test]
fn snapshot_scan_keeps_legacy_cell_reference_and_tag_semantics_together() {
    let cases: &[(&[u8], &str, bool)] = &[
        (br#"<c/>"#, "A7", true),
        (br#"<c><v>1</v></c>"#, "B7", false),
        (br#"<c r="C7"/>"#, "C7", true),
        (br#"<x:c r="D7"/>"#, "D7", true),
        (br#"<c r="E&#x37;"/>"#, "E7", true),
        (
            br#"<c z='before' r='F7' xmlns:x='urn:future' x:opaque='a&amp;b'/>"#,
            "F7",
            true,
        ),
    ];
    let cells_xml = cases
        .iter()
        .map(|(xml, _, _)| std::str::from_utf8(xml).expect("cell XML is UTF-8"))
        .collect::<String>();
    let source = format!(
        r#"<worksheet xmlns="{MAIN}" xmlns:x="{MAIN}"><sheetData><row r="7">{cells_xml}</row></sheetData></worksheet>"#
    );
    let layout = scan(source.as_bytes()).expect("worksheet scan");
    let cells = &layout.sheet_data.rows[0].cells;
    assert_eq!(cells.len(), cases.len());

    for (&(xml, expected_a1, expected_empty), cell) in cases.iter().zip(cells) {
        let (element, decoder) = start_element(xml);
        let reference = unqualified_attribute_value(&element, b"r", decoder)
            .expect("legacy cell reference lookup");
        if let Some(reference) = reference.as_deref() {
            let (reference_row, _) = parse_a1(reference).expect("legacy cell reference parse");
            assert_eq!(reference_row, 7);
        } else {
            assert!(matches!(expected_a1, "A7" | "B7"));
        }
        let expected_tag = cell_tag(&element, decoder).expect("legacy compact cell tag");

        assert_eq!(cell.address.a1(), expected_a1);
        assert_eq!(cell.empty, expected_empty);
        assert_eq!(cell.tag.is_some(), expected_tag.is_some());
        if let (Some(actual), Some(expected)) = (&cell.tag, &expected_tag) {
            assert_eq!(actual.name, expected.name);
            assert_eq!(actual.attributes.len(), expected.attributes.len());
            for (actual, expected) in actual.attributes.iter().zip(&expected.attributes) {
                assert_eq!(actual.name, expected.name);
                assert_eq!(actual.value, expected.value);
            }
        }
    }

    let prefixed = cells[3].tag.as_ref().expect("prefixed cell tag");
    assert_eq!(prefixed.name.as_ref(), "x:c");
    let retained = cells[5].tag.as_ref().expect("attribute-rich cell tag");
    assert_eq!(
        retained
            .attributes
            .iter()
            .map(|attribute| attribute.name.as_ref())
            .collect::<Vec<_>>(),
        ["z", "r", "xmlns:x", "x:opaque"]
    );
    assert_eq!(retained.attributes[0].value.as_ref(), "before");
    assert_eq!(retained.attributes[1].value.as_ref(), "F7");
    assert_eq!(retained.attributes[2].value.as_ref(), "urn:future");
    assert_eq!(retained.attributes[3].value.as_ref(), "a&b");
}

fn legacy_cell_pipeline_error(xml: &[u8]) -> (String, String) {
    let (element, decoder) = start_element(xml);
    let reference = match unqualified_attribute_value(&element, b"r", decoder) {
        Ok(reference) => reference,
        Err(error) => {
            let error = crate::error::Error::from(error);
            return (format!("{error:?}"), error.to_string());
        },
    };
    if let Some(reference) = reference {
        if let Err(error) = parse_a1(&reference) {
            return (format!("{error:?}"), error.to_string());
        }
    }
    if let Err(error) = cell_tag(&element, decoder) {
        return (format!("{error:?}"), error.to_string());
    }
    panic!("legacy cell pipeline unexpectedly accepted {xml:?}");
}

fn worksheet_with_cell(cell: &[u8]) -> Vec<u8> {
    let mut source = format!(r#"<worksheet xmlns="{MAIN}"><sheetData><row r="1">"#).into_bytes();
    source.extend_from_slice(cell);
    source.extend_from_slice(b"</row></sheetData></worksheet>");
    source
}

#[test]
fn snapshot_scan_preserves_cell_error_precedence_through_trailing_attributes() {
    let cases: &[&[u8]] = &[
        br#"<c r="A1" future="&missing;"/>"#,
        br#"<c r="A1" future="one" future="two"/>"#,
        br#"<c r="not-a-cell" future="&missing;"/>"#,
        b"<c r=\"\xff\" future=\"&missing;\"/>",
        b"<c r=\"A1\" future=\"\xff\"/>",
    ];

    for cell in cases {
        let expected = legacy_cell_pipeline_error(cell);
        let source = worksheet_with_cell(cell);
        let error = scan(&source).expect_err("cell pipeline should reject malformed input");
        assert_eq!(
            format!("{error:?}"),
            expected.0,
            "typed error changed for {cell:?}"
        );
        assert_eq!(
            error.to_string(),
            expected.1,
            "display error changed for {cell:?}"
        );
    }
}

#[test]
fn snapshot_scan_rejects_a_truncated_cell_start_after_reference_attributes() {
    let source =
        format!(r#"<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1" future="value""#)
            .into_bytes();
    let expected = {
        let mut reader = NsReader::from_reader(source.as_slice());
        reader.config_mut().check_end_names = true;
        loop {
            match reader.read_event() {
                Ok(Event::Eof) => panic!("truncated worksheet unexpectedly reached EOF"),
                Ok(_) => {},
                Err(error) => break crate::error::invalid(error.to_string()),
            }
        }
    };
    let error = scan(&source).expect_err("truncated cell start should be rejected");
    assert_eq!(format!("{error:?}"), format!("{expected:?}"));
    assert_eq!(error.to_string(), expected.to_string());
}

#[test]
fn snapshot_scan_reports_cell_address_errors_before_compact_tag_errors() {
    let duplicate_reference = format!(
        r#"<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1" r="B1" future="&missing;"/></row></sheetData></worksheet>"#
    );
    let error = scan(duplicate_reference.as_bytes())
        .expect_err("duplicate cell references should be rejected")
        .to_string();
    assert!(
        error.contains("duplicated attribute"),
        "unexpected error: {error}"
    );

    let mismatched_row = format!(
        r#"<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A2" future="&missing;"/></row></sheetData></worksheet>"#
    );
    let error = scan(mismatched_row.as_bytes())
        .expect_err("cell address must match its row")
        .to_string();
    assert!(
        error.contains("does not belong to row 1"),
        "unexpected error: {error}"
    );

    let malformed_reference = format!(
        r#"<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="&missing;" future="&later;" future="two"/></row></sheetData></worksheet>"#
    );
    let error = scan(malformed_reference.as_bytes())
        .expect_err("malformed cell reference should be rejected")
        .to_string();
    assert!(error.contains("missing"), "unexpected error: {error}");
    assert!(
        !error.contains("duplicate XML attribute"),
        "unexpected later error: {error}"
    );
}

#[test]
fn snapshot_scan_preserves_malformed_error_order_after_prior_primary_spans() {
    let cases: &[&[u8]] = &[
        br#"<c r="A1" future="&missing;"/>"#,
        br#"<c r="A1" future="one" future="two"/>"#,
        br#"<c r="not-a-cell" future="&missing;"/>"#,
        b"<c r=\"\xff\" future=\"&missing;\"/>",
        b"<c r=\"A1\" future=\"\xff\"/>",
    ];

    for cell in cases {
        let expected = legacy_cell_pipeline_error(cell);
        let mut source = format!(
            r#"<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1"><v>prior-number</v><!--prior-sentinel--><f>prior-formula</f><is><t>prior-text</t></is></c>"#
        )
        .into_bytes();
        source.extend_from_slice(cell);
        source.extend_from_slice(b"</row></sheetData></worksheet>");
        let error = scan(&source).expect_err("malformed cell should be rejected");
        assert_eq!(
            format!("{error:?}"),
            expected.0,
            "typed error changed after prior primary spans for {cell:?}"
        );
        assert_eq!(
            error.to_string(),
            expected.1,
            "display error changed after prior primary spans for {cell:?}"
        );
    }
}

fn assert_tag_results_match(xml: &[u8]) {
    let (old_element, old_decoder) = start_element(xml);
    let old_result = owned_tag(&old_element, old_decoder);
    let (new_element, new_decoder) = start_element(xml);
    let new_result = cell_tag(&new_element, new_decoder);
    match (old_result, new_result) {
        (Ok(old), Ok(Some(new))) => {
            assert_eq!(format!("{old:?}"), format!("{new:?}"));
        },
        (Err(old_error), Err(new_error)) => {
            assert_eq!(
                format!("{old_error:?}"),
                format!("{new_error:?}"),
                "debug error changed for {xml:?}"
            );
            assert_eq!(
                old_error.to_string(),
                new_error.to_string(),
                "display error changed for {xml:?}"
            );
        },
        (old, new) => panic!("owned and compact results diverged for {xml:?}: {old:?} vs {new:?}"),
    }
}

fn assert_tag_errors_match(xml: &[u8]) {
    let (old_element, old_decoder) = start_element(xml);
    let old_error = owned_tag(&old_element, old_decoder).expect_err("owned tag should fail");
    let (new_element, new_decoder) = start_element(xml);
    let new_error = cell_tag(&new_element, new_decoder).expect_err("compact tag should fail");
    assert_eq!(
        format!("{old_error:?}"),
        format!("{new_error:?}"),
        "debug error changed for {xml:?}"
    );
    assert_eq!(
        old_error.to_string(),
        new_error.to_string(),
        "display error changed for {xml:?}"
    );
}

#[test]
fn compact_cell_tag_keeps_attribute_error_order_and_messages() {
    let duplicate_cases: &[&[u8]] = &[
        br#"<c future="one" future="two" r="A1"/>"#,
        br#"<c r="A1" future="one" future="two"/>"#,
        br#"<c r="A1" r="A2"/>"#,
    ];
    for xml in duplicate_cases {
        assert_tag_errors_match(xml);
    }

    let malformed_cases: &[&[u8]] = &[
        b"<c \xff=\"value\"/>",
        b"<c r=\"\xff\"/>",
        br#"<c future="&missing;" later="value"/>"#,
        br#"<c r="A1" future="&missing;" later="value"/>"#,
        br#"<c future="&missing;" future="two"/>"#,
        br#"<c future="one" future="&missing;"/>"#,
        br#"<c r="A1" future="&missing;" future="two"/>"#,
    ];
    for xml in malformed_cases {
        assert_tag_results_match(xml);
    }
}
