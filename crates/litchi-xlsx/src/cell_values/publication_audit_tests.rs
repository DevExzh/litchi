#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "each test is one constructed witness with a fixed expected verdict"
)]
//! Witnesses for change 0747: neither publication audit of a replaced XML
//! Part is a duplicate of a check that runs earlier in the same operation.
//!
//! A source-backed value edit is admitted by planning (the value-only
//! validator with the raw worksheet parser, and the same validator with the
//! workbook catalog parser), rewritten by the commit, and then audited by
//! `litchi-opc` at publication with `xml_minifier::audit::verify_source`
//! under `Limits::default()`, once over each replaced Part's original bytes
//! and once over its replacement. Change 0747 asked whether planning had
//! already proved the original audit's verdict over the identical bytes, so
//! that publication could carry the proof instead of repeating the audit. It
//! had not. Every test below is an input planning and the commit admit and
//! publication refuses:
//!
//! * two refusals come from the audit of the *original* bytes alone, because
//!   the defect sits in the span the rewrite replaces: eliding that audit
//!   would publish them;
//! * three show a check the audit makes and planning does not, on bytes the
//!   rewrite copies, so both audits would refuse them;
//! * one comes from the audit of the *replacement* alone: the writer turns an
//!   original the audit accepts into bytes it refuses.
//!
//! Every refusal is asserted with its exact part, its exact audit diagnostic
//! and an empty sink: publication still fails closed before its first byte.

use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_opc::OpcError;
use litchi_opc::constants::content_type as ct;
use soapberry_zip::office::StreamingArchiveWriter;

use crate::cell_values::{SheetCellValueEdit, SourceBackedEditor};
use crate::{Address, Error, Number, Value};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const WORKSHEET: &str = "/xl/worksheets/sheet1.xml";

/// `xml_minifier::audit::Limits::default()`'s aggregate attribute budget,
/// which publication applies to both sides of a replacement.
const AUDIT_ATTRIBUTE_BUDGET: usize = 250_000;

/// Attribute names for the padding elements: 52 distinct one-letter names
/// keep every padding tag well below any per-element bound and keep the
/// duplicate-name checks of both scanners cheap.
const PAD_NAMES: &[u8] = b"abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";

fn package(worksheet: &str) -> Vec<u8> {
    package_with(&workbook(""), worksheet)
}

/// The one-sheet workbook, with `after_root` appended after its root.
fn workbook(after_root: &str) -> String {
    format!(
        "<workbook xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rIdSheet\"/></sheets></workbook>{after_root}"
    )
}

fn package_with(workbook: &str, worksheet: &str) -> Vec<u8> {
    let content_types = format!(
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"{}\"/><Override PartName=\"{WORKSHEET}\" ContentType=\"{}\"/></Types>",
        ct::SML_SHEET_MAIN,
        ct::SML_WORKSHEET,
    );
    let package_relationships = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"{PKG_REL}\"><Relationship Id=\"rIdRoot\" Type=\"{REL}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"
    );
    let workbook_relationships = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"{PKG_REL}\"><Relationship Id=\"rIdSheet\" Type=\"{REL}/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>"
    );
    let mut writer = StreamingArchiveWriter::new();
    for (name, bytes) in [
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", package_relationships.as_bytes()),
        ("xl/workbook.xml", workbook.as_bytes()),
        (
            "xl/_rels/workbook.xml.rels",
            workbook_relationships.as_bytes(),
        ),
        ("xl/worksheets/sheet1.xml", worksheet.as_bytes()),
    ] {
        writer.write_stored(name, bytes).expect("test member");
    }
    writer.finish_to_bytes().expect("test XLSX archive")
}

/// A one-row worksheet whose `A1` record is `a1` verbatim, followed by
/// `B1` and then `tail` after `</sheetData>`. Its fixed markup carries six
/// attributes plus whatever `a1` and `tail` add.
fn worksheet(a1: &str, tail: &str) -> String {
    format!(
        "<worksheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><dimension ref=\"A1:B1\"/><sheetData><row r=\"1\">{a1}<c r=\"B1\"><v>2</v></c></row></sheetData>{tail}</worksheet>"
    )
}

/// Attributes on the fixed markup of [`worksheet`] with a plain `A1` record:
/// the two namespace declarations, `dimension/@ref`, `row/@r` and the two
/// `c/@r`.
const FIXED_ATTRIBUTES: usize = 6;

/// A copied `<extLst>` payload carrying exactly `attributes` attributes,
/// counting the two on its `<ext>` wrapper.
fn attribute_pad(attributes: usize) -> String {
    let mut remaining = attributes.checked_sub(2).expect("the wrapper's two");
    let mut pad = String::from("<extLst><ext uri=\"{0747}\" xmlns:p=\"urn:litchi:0747:pad\">");
    while remaining > 0 {
        let count = remaining.min(PAD_NAMES.len());
        pad.push_str("<p:a");
        for name in &PAD_NAMES[..count] {
            pad.push(' ');
            pad.push(char::from(*name));
            pad.push_str("=\"\"");
        }
        pad.push_str("/>");
        remaining -= count;
    }
    pad.push_str("</ext></extLst>");
    pad
}

fn editor(package: Vec<u8>) -> SourceBackedEditor {
    SourceBackedEditor::from_read_at(Arc::new(OwnedSource::new(package)))
        .expect("the planner opens the witness package")
}

fn number(value: &str) -> Value {
    Value::Number(Number::new(value).expect("numeral"))
}

/// Plan, stage and commit one edit through the multi-sheet door the
/// measured edit/save route uses, then publish it to a fresh sink.
///
/// Returns the committed worksheet bytes, the publication result and the
/// sink. Planning and commit must succeed: every witness is an input the
/// planner admits.
fn plan_commit_publish(
    worksheet: &str,
    edit: SheetCellValueEdit<'_>,
) -> (Vec<u8>, Result<(), Error>, Vec<u8>) {
    plan_commit_publish_package(package(worksheet), edit)
}

fn plan_commit_publish_package(
    package: Vec<u8>,
    edit: SheetCellValueEdit<'_>,
) -> (Vec<u8>, Result<(), Error>, Vec<u8>) {
    let editor = editor(package);
    let commit = editor
        .edit_many([edit])
        .expect("planning admits the witness")
        .commit()
        .expect("the commit admits the witness");
    let replacement = commit
        .snapshot()
        .sheets()
        .first()
        .expect("one worksheet")
        .source_xml()
        .to_vec();
    let mut sink = Vec::new();
    let published = editor
        .publish_multi_commit_to_stream(&mut sink, &commit)
        .map(|_| ());
    (replacement, published, sink)
}

fn set_a1(value: &str) -> SheetCellValueEdit<'static> {
    SheetCellValueEdit::set("Sheet1", Address::from_a1("A1").expect("A1"), number(value))
}

/// Assert a publication refusal from the XML audit of `expected_part`, with
/// its exact diagnostic and no archive byte emitted.
fn assert_audit_refusal(
    published: Result<(), Error>,
    sink: &[u8],
    expected_part: &str,
    diagnostic: &str,
) {
    match published {
        Err(Error::Package(OpcError::XmlPublication { part, source })) => {
            assert_eq!(part, expected_part, "the refusal names the audited Part");
            assert_eq!(source.to_string(), diagnostic);
        },
        other => panic!("expected an XML audit refusal of {expected_part}, got {other:?}"),
    }
    assert!(
        sink.is_empty(),
        "a refused publication emits no archive byte"
    );
}

fn assert_worksheet_audit_refusal(published: Result<(), Error>, sink: &[u8], diagnostic: &str) {
    assert_audit_refusal(published, sink, WORKSHEET, diagnostic);
}

/// The replacement bytes pass the same audit when they are someone's
/// original: a package holding them publishes an ordinary edit.
fn assert_publishable(worksheet: &str) {
    let (_, published, sink) = plan_commit_publish(worksheet, set_a1("43"));
    published.expect("a package holding these bytes publishes");
    assert!(!sink.is_empty(), "the publication emitted an archive");
}

// ------------------------------------------------ the original audit

#[test]
fn an_invalid_space_value_in_the_replaced_value_element_is_refused_only_by_the_original_audit() {
    // The raw parser reads no attribute of `<v>`, and the value-only
    // validator refuses only a relationship reference inside `<sheetData>`,
    // so planning admits `xml:space="bogus"`. The rewrite drops the edited
    // cell's prior `<v>` span, so the replacement does not carry it.
    let a1 = "<c r=\"A1\"><v xml:space=\"bogus\">1</v></c>";
    let source = worksheet(a1, "");
    let offset = source.find("<v xml:space").expect("fixture offset");
    let (replacement, published, sink) = plan_commit_publish(&source, set_a1("42"));
    assert_worksheet_audit_refusal(
        published,
        &sink,
        &format!("malformed XML at byte {offset}: xml:space must be 'default' or 'preserve'"),
    );
    let replacement = String::from_utf8(replacement).expect("UTF-8 replacement");
    assert!(
        !replacement.contains("xml:space"),
        "the refused attribute is not in the replacement"
    );
    assert!(replacement.contains("<c r=\"A1\"><v>42</v></c>"));
    assert_publishable(&replacement);
}

#[test]
fn an_attribute_budget_exceeded_only_by_the_replaced_span_is_refused_only_by_the_original_audit() {
    // The planner has no aggregate attribute budget. One attribute over the
    // audit's budget sits on the edited cell's `<v>`, which the rewrite
    // replaces, so the original is one over the budget and the replacement
    // is exactly on it.
    let a1 = "<c r=\"A1\"><v xml:space=\"preserve\">1</v></c>";
    let pad = attribute_pad(AUDIT_ATTRIBUTE_BUDGET - FIXED_ATTRIBUTES);
    let source = worksheet(a1, &pad);
    let last_pad_tag = source.rfind("<p:a").expect("fixture offset");
    let (replacement, published, sink) = plan_commit_publish(&source, set_a1("42"));
    assert_worksheet_audit_refusal(
        published,
        &sink,
        &format!(
            "XML Attributes limit {AUDIT_ATTRIBUTE_BUDGET} exceeded by {} at byte {last_pad_tag}",
            AUDIT_ATTRIBUTE_BUDGET + 1
        ),
    );
    let replacement = String::from_utf8(replacement).expect("UTF-8 replacement");
    assert!(replacement.contains("<c r=\"A1\"><v>42</v></c>"));
    assert_publishable(&replacement);
}

#[test]
fn whitespace_cdata_outside_the_root_is_admitted_by_planning_and_refused_by_the_audit() {
    // The value-only validator admits whitespace-only character data outside
    // the root, CDATA included; the audit refuses any CDATA there. The tail
    // is copied, so the replacement carries it as well: this witnesses only
    // that planning's check is weaker than the audit, not which audit fires.
    let source = format!("{}<![CDATA[ ]]>", worksheet("<c r=\"A1\"><v>1</v></c>", ""));
    let offset = source.find("<![CDATA[").expect("fixture offset");
    let (_, published, sink) = plan_commit_publish(&source, set_a1("42"));
    assert_worksheet_audit_refusal(
        published,
        &sink,
        &format!("malformed XML at byte {offset}: CDATA outside the document element"),
    );
}

#[test]
fn whitespace_cdata_outside_the_workbook_root_is_admitted_by_planning_and_refused_by_the_audit() {
    // The workbook side of the same gap. Planning checks the workbook with
    // the same value-only validator and the catalog parser; the
    // calculation-properties splice copies every byte after `</workbook>`.
    let workbook = workbook("<![CDATA[ ]]>");
    let offset = workbook.find("<![CDATA[").expect("fixture offset");
    let package = package_with(&workbook, &worksheet("<c r=\"A1\"><v>1</v></c>", ""));
    let (_, published, sink) = plan_commit_publish_package(package, set_a1("42"));
    assert_audit_refusal(
        published,
        &sink,
        "/xl/workbook.xml",
        &format!("malformed XML at byte {offset}: CDATA outside the document element"),
    );
}

#[test]
fn an_attribute_with_no_separator_is_admitted_by_planning_and_refused_by_the_audit() {
    // quick-xml's checked attribute iterator, which the value-only validator
    // uses, accepts `a="1"b="2"`; the audit's attribute grammar refuses the
    // missing separator even under the source policy.
    let tail = "<extLst><ext uri=\"{0747}\" xmlns:p=\"urn:litchi:0747:pad\"><p:a a=\"1\"b=\"2\"/></ext></extLst>";
    let source = worksheet("<c r=\"A1\"><v>1</v></c>", tail);
    let tag = source.find("<p:a").expect("fixture offset");
    let separator = tag + "<p:a a=\"1\"".len();
    let (_, published, sink) = plan_commit_publish(&source, set_a1("42"));
    assert_worksheet_audit_refusal(
        published,
        &sink,
        &format!("malformed XML at byte {separator}: missing attribute separator"),
    );
}

// --------------------------------------------- the replacement audit

#[test]
fn an_insertion_can_push_an_admitted_original_over_the_attribute_budget() {
    // The original is exactly on the budget, so its audit passes: a
    // count-neutral edit of the same package publishes. An insertion adds
    // one `c/@r`, and only the replacement audit refuses the result. The
    // writer can therefore emit bytes the audit rejects from an original
    // the audit accepts, and the replacement audit is not redundant by
    // construction.
    let pad = attribute_pad(AUDIT_ATTRIBUTE_BUDGET - FIXED_ATTRIBUTES);
    let source = worksheet("<c r=\"A1\"><v>1</v></c>", &pad);

    let (_, published, sink) = plan_commit_publish(&source, set_a1("42"));
    published.expect("the original is within the audit budget");
    assert!(!sink.is_empty());

    let insert = SheetCellValueEdit::insert(
        "Sheet1",
        Address::from_a1("C1").expect("C1"),
        Number::new("7").expect("numeral"),
    );
    let (replacement, published, sink) = plan_commit_publish(&source, insert);
    let replacement = String::from_utf8(replacement).expect("UTF-8 replacement");
    assert!(replacement.contains("<c r=\"C1\"><v>7</v></c>"));
    let last_pad_tag = replacement.rfind("<p:a").expect("fixture offset");
    assert_worksheet_audit_refusal(
        published,
        &sink,
        &format!(
            "XML Attributes limit {AUDIT_ATTRIBUTE_BUDGET} exceeded by {} at byte {last_pad_tag}",
            AUDIT_ATTRIBUTE_BUDGET + 1
        ),
    );
}
