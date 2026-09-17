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
/// A refusal the value-only validator owns, placed after the padding so it
/// is reached only once the traversal has fallen back.
///
/// Change 0657 admitted `<mergeCells>`: it is outside `<sheetData>`, so the
/// rewrite copies it and the raw parser is the module that models it. An
/// unknown element *inside* a cell record is still the validator's to refuse,
/// because that span is the one the rewrite composes.
const LATE_VALIDATOR_TAIL: &str =
    "<sheetData><row r=\"1\"><c r=\"A1\"><future/><v>1</v></c></row></sheetData></worksheet>";
const LATE_RAW_TAIL: &str =
    "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v><v>2</v></c></row></sheetData></worksheet>";
const LATE_VALIDATOR_ERROR: &str =
    "value-only edits refuse dependency-bearing or unknown element 'future'";
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
    let config = reader.config_mut();
    config.allow_dangling_amp = false;
    config.allow_unmatched_ends = false;
    config.check_comments = false;
    config.check_end_names = true;
    config.expand_empty_elements = false;
    config.trim_markup_names_in_closing_tags = true;
    config.trim_text(false);
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
    assert!(!crate::raw::worksheet::shared_event_bound_within_cap(
        &sheet
    ));
    assert_valid_noop(&sheet);
}

#[test]
fn exact_text_cap_and_chunk_boundaries() {
    let cap = crate::raw::worksheet::MAX_SHARED_PROVISIONAL_EVENTS;
    let mut exact = b"<x/>a".repeat(65_535);
    exact.extend_from_slice(b"<x/>");
    assert_eq!(independent_lexical_event_bound(&exact), cap);
    assert_conservative_event_bound("exact-text-cap", &exact);
    assert!(crate::raw::worksheet::shared_event_bound_within_cap(&exact));

    exact.push(b'a');
    assert_eq!(independent_lexical_event_bound(&exact), cap + 1);
    assert_conservative_event_bound("exact-text-cap-plus-one", &exact);
    assert!(!crate::raw::worksheet::shared_event_bound_within_cap(
        &exact
    ));

    for offset in [4095, 4096, 8191, 8192, 65_535, 65_536] {
        for end in [b'>', b';'] {
            for next in [b'<', b'&', b'a'] {
                let mut data = vec![b' '; offset];
                data.extend_from_slice(&[end, next]);
                let count = cap - independent_lexical_event_bound(&data);
                data.extend_from_slice(&COMMENT.as_bytes().repeat(count));
                assert_eq!(independent_lexical_event_bound(&data), cap);
                assert_conservative_event_bound("chunk-boundary-exact", &data);
                assert!(crate::raw::worksheet::shared_event_bound_within_cap(&data));

                data.extend_from_slice(COMMENT.as_bytes());
                assert_conservative_event_bound("chunk-boundary-over", &data);
                assert!(!crate::raw::worksheet::shared_event_bound_within_cap(&data));
            }
        }
    }

    for final_byte in [b'<', b'&', b'>', b';'] {
        let data = [final_byte];
        assert_conservative_event_bound("final-byte", &data);
    }
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

// ---------------------------------------------------------------------------
// Markup-compatibility admission (change 0603)
// ---------------------------------------------------------------------------

const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const X14AC: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

/// The established two-pass path: complete value-only validation of the source,
/// then the preprocessing raw parse of the same source.
fn authoritative(content: &[u8]) -> crate::Result<crate::cell::Store> {
    super::worksheet_xml(content)?;
    crate::raw::worksheet::parse(content, || Ok(None))
}

/// The admitted path, exactly as `Snapshot::from_source_selected` drives it.
fn admitted(content: &[u8]) -> crate::Result<crate::cell::Store> {
    match crate::raw::worksheet::source_stream_admission(content) {
        Some(admission) => super::worksheet_xml_and_parse_source(content, admission, || Ok(None))
            .map(|(cells, _)| cells),
        None => authoritative(content),
    }
}

/// Run only the raw shared traversal and report whether the rewrite proof
/// reached EOF. The production validator is intentionally not part of this
/// probe: the real-part census uses it to separate the MCE proof from the
/// value-only vocabulary and from the authoritative fallback.
fn rewritten_proof_completes(content: &[u8]) -> bool {
    matches!(
        crate::raw::worksheet::parse_source_with_observer(
            content,
            crate::raw::worksheet::SourceAdmission::Rewritten,
            || Ok(None),
            |_, _, _| true,
        ),
        crate::raw::worksheet::SourceParseAttempt::Complete(_)
    )
}

/// Report whether the complete value-only admission kept the source-backed
/// facts from the shared traversal. `None` means that the validator or the
/// rewrite proof deliberately selected the established fallback.
fn shared_source_completes(content: &[u8]) -> bool {
    let Some(admission) = crate::raw::worksheet::source_stream_admission(content) else {
        return false;
    };
    let mut validator = super::Validator::new(super::XmlOwner::Worksheet);
    let attempt = crate::raw::worksheet::parse_source_with_observer(
        content,
        admission,
        || Ok(None),
        |namespace, event, _| validator.observe(namespace, event),
    );
    matches!(
        attempt,
        crate::raw::worksheet::SourceParseAttempt::Complete(_)
    ) && validator.finish().is_ok()
}

/// Report whether admission changed the parsed value or the exact error.
fn admission_difference(name: &str, content: &[u8]) -> Option<String> {
    match (authoritative(content), admitted(content)) {
        (Ok(expected), Ok(actual)) => {
            let (expected, actual) = (format!("{expected:?}"), format!("{actual:?}"));
            (expected != actual).then(|| format!("{name}: admitted store differs"))
        },
        (Err(expected), Err(actual)) => {
            let (expected, actual) = (format!("{expected}"), format!("{actual}"));
            (expected != actual)
                .then(|| format!("{name}: admitted error '{actual}' replaces '{expected}'"))
        },
        (Ok(_), Err(actual)) => Some(format!(
            "{name}: admission refused an accepted worksheet: {actual}"
        )),
        (Err(expected), Ok(_)) => Some(format!(
            "{name}: admission accepted a worksheet refused with '{expected}'"
        )),
    }
}

/// Assert that admission changed neither the parsed value nor the exact error.
fn assert_admission_is_transparent(name: &str, content: &[u8]) {
    assert_eq!(admission_difference(name, content), None);
}

fn worksheet_with_root_attributes(attributes: &str, body: &str) -> Vec<u8> {
    format!("<worksheet xmlns=\"{SML}\"{attributes}>{body}</worksheet>").into_bytes()
}

#[test]
fn declaration_only_markers_reach_the_shared_traversal() {
    use crate::raw::worksheet::{SourceAdmission, source_stream_admission};

    let declaration_only = worksheet_with_root_attributes(
        &format!(" xmlns:mc=\"{MCE}\" xmlns:x14ac=\"{X14AC}\""),
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>",
    );
    assert_eq!(
        source_stream_admission(&declaration_only),
        Some(SourceAdmission::Rewritten)
    );
    assert!(rewritten_proof_completes(&declaration_only));
    assert_valid_noop(&declaration_only);
    assert_admission_is_transparent("declaration-only", &declaration_only);

    // An x14ac declaration without a descent value never reached the shared
    // reader before; with no markup-compatibility namespace the preprocessor
    // borrows, so the traversal needs no proof at all.
    let extension_only = worksheet_with_root_attributes(
        &format!(" xmlns:x14ac=\"{X14AC}\""),
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>",
    );
    assert_eq!(
        source_stream_admission(&extension_only),
        Some(SourceAdmission::Borrowed)
    );
    assert_valid_noop(&extension_only);

    // The two markers that still refuse admission outright.
    let descent = worksheet_with_root_attributes(
        &format!(" xmlns:mc=\"{MCE}\" xmlns:x14ac=\"{X14AC}\" mc:Ignorable=\"x14ac\""),
        "<sheetData><row r=\"1\" x14ac:dyDescent=\"0.25\"><c r=\"A1\"><v>1</v></c></row></sheetData>",
    );
    assert_eq!(source_stream_admission(&descent), None);
    let alternate = worksheet_with_root_attributes(
        &format!(" xmlns:mc=\"{MCE}\""),
        "<sheetData/><mc:AlternateContent><mc:Choice Requires=\"x\"/></mc:AlternateContent>",
    );
    assert_eq!(source_stream_admission(&alternate), None);
}

#[test]
fn rewrite_declaration_budget_is_per_start_tag() {
    let levels = 96usize;
    let declarations_per_level = 3usize;
    let mut xml = format!("<worksheet xmlns=\"{SML}\" xmlns:mc=\"{MCE}\">");
    for level in 0..levels {
        write!(xml, "<extension{level}").expect("write extension start");
        for declaration in 0..declarations_per_level {
            write!(
                xml,
                " xmlns:p{level}_{declaration}=\"urn:litchi:{level}:{declaration}\""
            )
            .expect("write extension namespace");
        }
        xml.push('>');
    }
    for level in (0..levels).rev() {
        write!(xml, "</extension{level}>").expect("write extension end");
    }
    xml.push_str("<sheetData/></worksheet>");
    let xml = xml.into_bytes();
    assert_eq!(
        crate::raw::worksheet::source_stream_admission(&xml),
        Some(crate::raw::worksheet::SourceAdmission::Rewritten)
    );
    assert!(
        levels * declarations_per_level > crate::raw::worksheet::MAX_REWRITTEN_DECLARATIONS,
        "fixture must exceed the old cumulative declaration budget"
    );
    assert!(rewritten_proof_completes(&xml));
}

#[test]
fn ignorable_directive_proof_is_narrow_and_transparent() {
    let tail = "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>";
    let valid = worksheet_with_root_attributes(
        &format!(
            " xmlns:mc=\"{MCE}\" xmlns:future=\"urn:litchi:future\" xmlns:other=\"urn:litchi:other\" mc:Ignorable=\"future\""
        ),
        &format!(
            "<sheetData future:flag=\"ignored\" other:flag=\"retained\"><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>"
        ),
    );
    assert_eq!(
        crate::raw::worksheet::source_stream_admission(&valid),
        Some(crate::raw::worksheet::SourceAdmission::Rewritten)
    );
    assert!(rewritten_proof_completes(&valid));
    assert!(shared_source_completes(&valid));
    assert_valid_noop(&valid);
    assert_admission_is_transparent("valid Ignorable", &valid);

    let unbound = worksheet_with_root_attributes(
        &format!(" xmlns:mc=\"{MCE}\" mc:Ignorable=\"missing\""),
        tail,
    );
    assert!(!rewritten_proof_completes(&unbound));
    assert_admission_is_transparent("unbound Ignorable", &unbound);

    let recursive =
        worksheet_with_root_attributes(&format!(" xmlns:mc=\"{MCE}\" mc:Ignorable=\"mc\""), tail);
    assert!(!rewritten_proof_completes(&recursive));
    assert_admission_is_transparent("MCE Ignorable", &recursive);

    let duplicate = worksheet_with_root_attributes(
        &format!(
            " xmlns:mc=\"{MCE}\" xmlns:future=\"urn:litchi:future\" mc:Ignorable=\"future future\""
        ),
        tail,
    );
    assert!(!rewritten_proof_completes(&duplicate));
    assert_admission_is_transparent("duplicate Ignorable", &duplicate);

    let process = worksheet_with_root_attributes(
        &format!(
            " xmlns:mc=\"{MCE}\" xmlns:future=\"urn:litchi:future\" mc:Ignorable=\"future\" mc:ProcessContent=\"future:future\""
        ),
        &format!("<future:future><marker/></future:future>{tail}"),
    );
    assert!(!rewritten_proof_completes(&process));
    assert_admission_is_transparent("ProcessContent", &process);

    let preserve = worksheet_with_root_attributes(
        &format!(
            " xmlns:mc=\"{MCE}\" xmlns:future=\"urn:litchi:future\" mc:Ignorable=\"future\" mc:PreserveAttributes=\"future:flag\""
        ),
        &format!("<future:future future:flag=\"kept\"/>{tail}"),
    );
    assert!(!rewritten_proof_completes(&preserve));
    assert_admission_is_transparent("PreserveAttributes", &preserve);
}

#[test]
fn rewrite_only_refusals_survive_marker_admission() {
    let tail = "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>";
    let declared = format!(" xmlns:mc=\"{MCE}\"");
    let mut cases: Vec<(&str, Vec<u8>)> = vec![
        // The preprocessor rejects processing instructions and DTDs outright,
        // where the worksheet parser ignores both.
        (
            "processing instruction",
            worksheet_with_root_attributes(&declared, &format!("<?target data?>{tail}")),
        ),
        (
            "leading processing instruction",
            format!(
                "<?target data?><worksheet xmlns=\"{SML}\" xmlns:mc=\"{MCE}\">{tail}</worksheet>"
            )
            .into_bytes(),
        ),
        // A declaration inside the root is rejected as late.
        (
            "late declaration",
            worksheet_with_root_attributes(&declared, &format!("<?xml version=\"1.0\"?>{tail}")),
        ),
        // Only the predefined entities and character references are re-emitted.
        (
            "custom entity",
            worksheet_with_root_attributes(
                &declared,
                "<sheetData><row r=\"1\"><c r=\"A1\"><v>&custom;</v></c></row></sheetData>",
            ),
        ),
        // An empty namespace value is an undeclaration to the reader and an
        // invalid namespace to the preprocessor.
        (
            "empty namespace value",
            worksheet_with_root_attributes(&format!("{declared} xmlns:empty=\"\""), tail),
        ),
        // A name the reader accepts but the preprocessor refuses as an
        // invalid QName.
        (
            "invalid element name",
            worksheet_with_root_attributes(&declared, &format!("{tail}<1bad/>")),
        ),
        (
            "invalid attribute name",
            worksheet_with_root_attributes(
                &declared,
                "<sheetData><row r=\"1\" 1bad=\"2\"><c r=\"A1\"><v>1</v></c></row></sheetData>",
            ),
        ),
        // A prefixed element may be unwrapped, skipped or refused.
        (
            "prefixed element",
            worksheet_with_root_attributes(
                &format!("{declared} xmlns:u=\"urn:litchi:unknown\""),
                &format!("{tail}<u:future/>"),
            ),
        ),
        // An unbound prefix parses and does not preprocess.
        (
            "unbound attribute prefix",
            worksheet_with_root_attributes(
                &declared,
                &format!(
                    "<sheetData><row r=\"1\" u:flag=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>{}",
                    ""
                ),
            ),
        ),
        // An undecodable attribute value is read by the preprocessor and not by
        // the parser.
        (
            "unrecognized attribute entity",
            worksheet_with_root_attributes(
                &declared,
                "<sheetData><row r=\"1\" spans=\"&custom;\"><c r=\"A1\"><v>1</v></c></row></sheetData>",
            ),
        ),
    ];
    // More declarations than one rewritten start tag may carry.
    let mut flood = String::from(&declared);
    for index in 0..=crate::raw::worksheet::MAX_REWRITTEN_DECLARATIONS {
        let _ = write!(flood, " xmlns:p{index}=\"urn:litchi:{index}\"");
    }
    cases.push((
        "declaration flood",
        worksheet_with_root_attributes(&flood, tail),
    ));

    let differences: Vec<String> = cases
        .iter()
        .filter_map(|(name, sheet)| admission_difference(name, sheet))
        .collect();
    assert!(
        differences.is_empty(),
        "marker admission changed {} of {} rewrite-only outcomes: {differences:#?}",
        differences.len(),
        cases.len()
    );
}

#[test]
fn marker_admission_matches_the_authoritative_path_on_every_real_worksheet() {
    use litchi_opc::OpcPackage;

    let root = concat!(env!("CARGO_MANIFEST_DIR"), "/../../test-data/ooxml/xlsx");
    let mut paths: Vec<std::path::PathBuf> = std::fs::read_dir(root)
        .expect("fixture directory")
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .filter(|path| path.extension().is_some_and(|value| value == "xlsx"))
        .collect();
    paths.sort();
    assert!(
        paths.len() >= 90,
        "unexpected fixture count {}",
        paths.len()
    );

    let mut parts = 0usize;
    let mut rewritten = 0usize;
    let mut rewritten_shared = 0usize;
    let mut borrowed = 0usize;
    for path in &paths {
        let Ok(package) = OpcPackage::open(path) else {
            continue;
        };
        for part in package
            .try_iter_parts()
            .map(|part| part.expect("part payload decodes"))
        {
            let name = part.partname().as_str().to_owned();
            if !name.starts_with("/xl/worksheets/sheet") || !name.ends_with(".xml") {
                continue;
            }
            let content = part.blob();
            parts += 1;
            match crate::raw::worksheet::source_stream_admission(content) {
                Some(crate::raw::worksheet::SourceAdmission::Rewritten) => {
                    rewritten += 1;
                    rewritten_shared += usize::from(shared_source_completes(content));
                },
                Some(crate::raw::worksheet::SourceAdmission::Borrowed) => borrowed += 1,
                None => {},
            }
            assert_admission_is_transparent(&format!("{}{name}", path.display()), content);
        }
    }
    assert!(parts >= 200, "unexpected worksheet part count {parts}");
    assert!(borrowed > 0, "no worksheet took the borrowed traversal");
    assert!(
        rewritten > 0,
        "no real worksheet exercised the rewrite-equivalence proof"
    );
    println!(
        "worksheet MCE admission census: parts={parts} borrowed={borrowed} rewritten={rewritten} rewritten_shared_complete={rewritten_shared} fallback={}",
        parts - borrowed - rewritten
    );
    assert!(
        rewritten_shared > 0,
        "no real worksheet completed the updated shared admission"
    );
}
