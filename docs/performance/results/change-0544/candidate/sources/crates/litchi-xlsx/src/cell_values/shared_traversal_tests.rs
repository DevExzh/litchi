#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused shared-traversal boundary assertions"
)]

use std::{fmt::Write as _, sync::Arc};

use litchi_core::OwnedSource;
use litchi_opc::constants::content_type as ct;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use soapberry_zip::office::StreamingArchiveWriter;

use crate::cell_values::{SheetCellValueEdit, SourceBackedEditor};
use crate::{Address, Error, Number, Value};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const COMMENT: &str = "<!--x-->";
const VALID_TAIL: &str =
    "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>";
const LATE_VALIDATOR_TAIL: &str =
    "<mergeCells/><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>";
const LATE_RAW_TAIL: &str =
    "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v><v>2</v></c></row></sheetData></worksheet>";
const LATE_VALIDATOR_ERROR: &str =
    "value-only edits refuse dependency-bearing or unknown element 'mergeCells'";
const LATE_RAW_ERROR: &str = "duplicate worksheet cell value";
const DELIMITER_COMMENT: &str = "<!-->>>>>>>>>>;;;;;;;;;;-->";

#[derive(Debug, Eq, PartialEq)]
enum DirectEventCount {
    Complete(usize),
    ReaderError(usize),
}

impl DirectEventCount {
    fn emitted(&self) -> usize {
        match self {
            Self::Complete(count) | Self::ReaderError(count) => *count,
        }
    }
}

fn worksheet_with_comments(comment_count: usize, tail: &str) -> Vec<u8> {
    let mut xml = String::with_capacity(64 + comment_count * COMMENT.len() + tail.len());
    xml.push_str("<worksheet xmlns=\"");
    xml.push_str(SML);
    xml.push_str("\">");
    for _ in 0..comment_count {
        xml.push_str(COMMENT);
    }
    xml.push_str(tail);
    xml.into_bytes()
}

fn worksheet_over_byte_cap(tail: &str) -> Vec<u8> {
    let count = crate::raw::worksheet::MAX_SHARED_SOURCE_BYTES / COMMENT.len() + 1;
    let xml = worksheet_with_comments(count, tail);
    assert!(xml.len() > crate::raw::worksheet::MAX_SHARED_SOURCE_BYTES);
    assert!(!crate::raw::worksheet::source_stream_eligible(&xml));
    xml
}

fn worksheet_over_event_cap(tail: &str) -> Vec<u8> {
    let count = crate::raw::worksheet::MAX_SHARED_PROVISIONAL_EVENTS + 1;
    let xml = worksheet_with_comments(count, tail);
    assert!(xml.len() < crate::raw::worksheet::MAX_SHARED_SOURCE_BYTES);
    assert!(crate::raw::worksheet::source_stream_eligible(&xml));
    assert!(!crate::raw::worksheet::shared_event_bound_within_cap(&xml));
    xml
}

fn worksheet_with_delimiter_comments(comment_count: usize, tail: &str) -> Vec<u8> {
    let mut xml = String::with_capacity(64 + comment_count * DELIMITER_COMMENT.len() + tail.len());
    xml.push_str("<worksheet xmlns=\"");
    xml.push_str(SML);
    xml.push_str("\">");
    for _ in 0..comment_count {
        xml.push_str(DELIMITER_COMMENT);
    }
    xml.push_str(tail);
    xml.into_bytes()
}

fn direct_ns_reader_event_count(content: &[u8]) -> DirectEventCount {
    let mut reader = NsReader::from_reader(content);
    reader.config_mut().check_end_names = true;
    let mut count = 0usize;
    loop {
        let event = match reader.read_event() {
            Ok(event) => event,
            Err(_) => return DirectEventCount::ReaderError(count),
        };
        let eof = matches!(event, Event::Eof);
        let _ = reader.resolver().resolve_event(event);
        count += 1;
        if eof {
            return DirectEventCount::Complete(count);
        }
    }
}

fn independent_lexical_event_bound(content: &[u8]) -> usize {
    let mut bound = 1usize;
    if content
        .first()
        .is_some_and(|&first| !matches!(first, b'<' | b'&'))
    {
        bound += 1;
    }
    for &byte in content {
        if matches!(byte, b'<' | b'&') {
            bound += 1;
        }
    }
    for (index, &byte) in content.iter().enumerate() {
        if matches!(byte, b'>' | b';')
            && content
                .get(index + 1)
                .is_some_and(|&next| !matches!(next, b'<' | b'&'))
        {
            bound += 1;
        }
    }
    bound
}

fn assert_conservative_event_bound(name: &str, content: &[u8]) -> DirectEventCount {
    let direct = direct_ns_reader_event_count(content);
    let independent_bound = independent_lexical_event_bound(content);
    let bound_within_cap = crate::raw::worksheet::shared_event_bound_within_cap(content);
    assert_eq!(
        bound_within_cap,
        independent_bound <= crate::raw::worksheet::MAX_SHARED_PROVISIONAL_EVENTS,
        "{name}: production bound disagrees with independent bound {independent_bound}"
    );
    assert!(
        direct.emitted() <= independent_bound,
        "{name}: direct event prefix exceeded independent bound: {direct:?} > {independent_bound}"
    );
    if direct.emitted() > crate::raw::worksheet::MAX_SHARED_PROVISIONAL_EVENTS {
        assert!(
            !bound_within_cap,
            "{name}: lexical bound admitted an over-cap reader stream: {direct:?}"
        );
    }
    direct
}

fn worksheet_numeric_grid(side: usize) -> Vec<u8> {
    fn column_name(mut index: usize) -> String {
        let mut name = String::new();
        loop {
            name.push((b'A' + (index % 26) as u8) as char);
            if index < 26 {
                break;
            }
            index = index / 26 - 1;
        }
        name.chars().rev().collect()
    }

    let columns: Vec<String> = (0..side).map(column_name).collect();
    let mut xml = String::with_capacity(side * side * 32);
    xml.push_str("<worksheet xmlns=\"");
    xml.push_str(SML);
    xml.push_str("\"><sheetData>");
    for row in 1..=side {
        write!(xml, "<row r=\"{row}\">").expect("row XML");
        for column in &columns {
            write!(xml, "<c r=\"{column}{row}\"><v>1</v></c>").expect("cell XML");
        }
        xml.push_str("</row>");
    }
    xml.push_str("</sheetData></worksheet>");
    xml.into_bytes()
}

fn package_with_sheet(sheet: &[u8]) -> Vec<u8> {
    let content_types = format!(
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"{}\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"{}\"/></Types>",
        ct::SML_SHEET_MAIN,
        ct::SML_WORKSHEET,
    );
    let workbook = format!(
        "<workbook xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rIdSheet\"/></sheets></workbook>"
    );
    let workbook_relationships = format!(
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdSheet\" Type=\"{REL}/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>"
    );
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .expect("content-types member");
    writer
        .write_stored(
            "_rels/.rels",
            b"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdRoot\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>",
        )
        .expect("package relationships member");
    writer
        .write_stored("xl/workbook.xml", workbook.as_bytes())
        .expect("workbook member");
    writer
        .write_stored(
            "xl/_rels/workbook.xml.rels",
            workbook_relationships.as_bytes(),
        )
        .expect("workbook relationships member");
    writer
        .write_stored("xl/worksheets/sheet1.xml", sheet)
        .expect("worksheet member");
    writer.finish_to_bytes().expect("test XLSX archive")
}

fn source_and_editor(sheet: &[u8]) -> (Vec<u8>, Arc<OwnedSource>, SourceBackedEditor) {
    let bytes = package_with_sheet(sheet);
    let source = Arc::new(OwnedSource::new(bytes.clone()));
    let editor = SourceBackedEditor::from_read_at(source.clone()).expect("accepted XLSX source");
    (bytes, source, editor)
}

fn assert_invalid<T>(result: crate::Result<T>, expected: &str) {
    match result {
        Err(Error::Invalid(actual)) => assert_eq!(actual, expected),
        Err(other) => panic!("expected typed Invalid error, got {other:?}"),
        Ok(_) => panic!("expected source-backed edit to be rejected"),
    }
}

fn assert_valid_noop(sheet: &[u8]) {
    let (bytes, source, editor) = source_and_editor(sheet);
    let commit = editor
        .edit_sheets(["Sheet1".into()])
        .expect("cap fallback should accept valid worksheet")
        .commit()
        .expect("valid no-op commit");
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(
        commit
            .snapshot()
            .value(0, Address::from_a1("A1").expect("A1 address")),
        Some(&Value::Number(Number::new("1").expect("source number")))
    );
    assert_eq!(source.as_slice(), bytes.as_slice());
}

fn assert_rejected_with_retry(sheet: &[u8], expected: &str) {
    let (bytes, source, editor) = source_and_editor(sheet);
    for attempt in 0..2 {
        assert_invalid(editor.edit_sheets(["Sheet1".into()]), expected);
        assert_eq!(
            source.as_slice(),
            bytes.as_slice(),
            "source changed after rejected attempt {attempt}"
        );
    }
}

#[test]
fn source_size_cap_falls_back_without_refusing_valid_or_late_failures() {
    assert_valid_noop(&worksheet_over_byte_cap(VALID_TAIL));
    assert_rejected_with_retry(
        &worksheet_over_byte_cap(LATE_VALIDATOR_TAIL),
        LATE_VALIDATOR_ERROR,
    );
    assert_rejected_with_retry(&worksheet_over_byte_cap(LATE_RAW_TAIL), LATE_RAW_ERROR);
}

#[test]
fn provisional_event_cap_falls_back_without_refusing_valid_or_late_failures() {
    assert_valid_noop(&worksheet_over_event_cap(VALID_TAIL));
    assert_rejected_with_retry(
        &worksheet_over_event_cap(LATE_VALIDATOR_TAIL),
        LATE_VALIDATOR_ERROR,
    );
    assert_rejected_with_retry(&worksheet_over_event_cap(LATE_RAW_TAIL), LATE_RAW_ERROR);
}

#[test]
fn event_bound_admits_main_numeric_grids() {
    for side in [96, 128] {
        let xml = worksheet_numeric_grid(side);
        assert!(xml.len() < crate::raw::worksheet::MAX_SHARED_SOURCE_BYTES);
        assert!(crate::raw::worksheet::source_stream_eligible(&xml));
        assert!(crate::raw::worksheet::shared_event_bound_within_cap(&xml));
    }
}

#[test]
fn lexical_event_bound_matches_direct_ns_reader_edge_oracle() {
    let mut bom = vec![0xef, 0xbb, 0xbf];
    bom.extend_from_slice(b"<a/>");
    let cases = vec![
        ("empty", b"".to_vec(), true),
        ("whitespace", b" \t\r\n".to_vec(), true),
        ("bom", bom, true),
        ("declaration", b"<?xml version=\"1.0\"?><a/>".to_vec(), true),
        ("processing_instruction", b"<?pi data?><a/>".to_vec(), true),
        ("comment", b"<!--comment--><a/>".to_vec(), true),
        (
            "comment_embedded_delimiters",
            b"<!-- > & ; --><a/>".to_vec(),
            true,
        ),
        ("cdata", b"<![CDATA[data]]><a/>".to_vec(), true),
        (
            "cdata_embedded_delimiters",
            b"<![CDATA[ > & ; ]]> <a/>".to_vec(),
            true,
        ),
        ("doctype", b"<!DOCTYPE a><a/>".to_vec(), true),
        (
            "doctype_embedded_delimiters",
            b"<!DOCTYPE a [ <!ELEMENT a ANY> <!ENTITY x \"a > & ;\"> ]><a/>".to_vec(),
            true,
        ),
        (
            "attribute_and_text_delimiters",
            b"<a attr=\"> &amp;\">text>;></a>".to_vec(),
            true,
        ),
        ("ordinary_text_delimiters", b"<a>text>;></a>".to_vec(), true),
        ("named_reference", b"<a>x&amp;y</a>".to_vec(), true),
        ("numeric_reference", b"<a>x&#65;y</a>".to_vec(), true),
        ("hex_numeric_reference", b"<a>x&#x41;y</a>".to_vec(), true),
        ("reference_chain", b"<a>&amp;&#65;</a>".to_vec(), true),
        ("dangling_reference", b"<a>x&broken</a>".to_vec(), false),
        (
            "reference_before_markup",
            b"<a>x&broken<b/></a>".to_vec(),
            false,
        ),
        (
            "reference_before_reference",
            b"<a>x&broken&ok;</a>".to_vec(),
            false,
        ),
        (
            "unterminated_numeric_reference",
            b"<a>&#65</a>".to_vec(),
            false,
        ),
        ("malformed_reference_body", b"<a>&#;</a>".to_vec(), true),
        ("malformed_nested", b"<a><b></a>".to_vec(), false),
        ("malformed_tail", b"<a><".to_vec(), false),
        ("nested_and_empty", b"<a><b/><c>v</c></a>".to_vec(), true),
    ];

    for (name, content, expected_complete) in cases {
        let direct = assert_conservative_event_bound(name, &content);
        assert_eq!(
            matches!(&direct, DirectEventCount::Complete(_)),
            expected_complete,
            "{name}: direct NsReader outcome was {direct:?}"
        );
    }

    let base = worksheet_with_comments(0, VALID_TAIL);
    let base_events = match direct_ns_reader_event_count(&base) {
        DirectEventCount::Complete(count) => count,
        other => panic!("base cap fixture did not reach EOF: {other:?}"),
    };
    let cap = crate::raw::worksheet::MAX_SHARED_PROVISIONAL_EVENTS;
    assert!(base_events < cap);
    assert_eq!(cap, 131_072);

    let exact = worksheet_with_comments(cap - base_events, VALID_TAIL);
    assert_eq!(independent_lexical_event_bound(&exact), cap);
    assert_eq!(
        direct_ns_reader_event_count(&exact),
        DirectEventCount::Complete(cap)
    );
    assert!(crate::raw::worksheet::shared_event_bound_within_cap(&exact));

    let over = worksheet_with_comments(cap - base_events + 1, VALID_TAIL);
    assert_eq!(independent_lexical_event_bound(&over), 131_073);
    assert_eq!(
        direct_ns_reader_event_count(&over),
        DirectEventCount::Complete(131_073)
    );
    assert!(!crate::raw::worksheet::shared_event_bound_within_cap(&over));
}

#[test]
fn lexical_event_bound_false_positive_still_uses_authoritative_fallback() {
    let sheet = worksheet_with_delimiter_comments(7_000, VALID_TAIL);
    let direct = direct_ns_reader_event_count(&sheet);
    let independent_bound = independent_lexical_event_bound(&sheet);
    assert!(matches!(&direct, DirectEventCount::Complete(_)));
    assert!(direct.emitted() < crate::raw::worksheet::MAX_SHARED_PROVISIONAL_EVENTS);
    assert!(independent_bound > crate::raw::worksheet::MAX_SHARED_PROVISIONAL_EVENTS);
    assert!(direct.emitted() < independent_bound);
    assert!(crate::raw::worksheet::source_stream_eligible(&sheet));
    assert!(!crate::raw::worksheet::shared_event_bound_within_cap(&sheet));
    assert_valid_noop(&sheet);
}

#[test]
fn eligible_small_source_uses_successful_value_edit_control() {
    let sheet = worksheet_with_comments(3, VALID_TAIL);
    assert!(crate::raw::worksheet::source_stream_eligible(&sheet));
    let (bytes, source, editor) = source_and_editor(&sheet);
    let mut edit = editor
        .edit_sheets(["Sheet1".into()])
        .expect("eligible worksheet");
    edit.apply_batch([SheetCellValueEdit::set(
        "Sheet1",
        Address::from_a1("A1").expect("A1 address"),
        2u32,
    )])
    .expect("stage value edit");
    let commit = edit.commit().expect("commit value edit");
    assert!(commit.changed());
    assert_eq!(
        commit
            .snapshot()
            .value(0, Address::from_a1("A1").expect("A1 address")),
        Some(&Value::Number(Number::new("2").expect("edited number")))
    );
    assert_eq!(source.as_slice(), bytes.as_slice());
}
