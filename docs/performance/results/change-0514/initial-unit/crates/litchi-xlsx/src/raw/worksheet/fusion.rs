//! Differential coverage for the conditional semantic/snapshot worksheet pass.
//!
//! These tests live below `raw::worksheet::tests`, so they can inspect the
//! transaction-local deferred layout without widening the production API.

use std::collections::BTreeMap;
use std::sync::Arc;

use litchi_sheet::Cell as Address;

use super::super::edit::{Action, Plan};
use super::super::{parse, parse_with_layout, x14ac};
use crate::cell::{Cell, Content, Store};

const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn source_fixture() -> Arc<Vec<u8>> {
    Arc::new(
        format!(
            r#"<worksheet xmlns="{S}" xmlns:x="{S}" xmlns:q="urn:litchi:future" q:root="keep"><dimension ref="A1:C2"/><sheetData q:container="keep"><row r="1" q:row="keep"><c r="A1" q:opaque="yes"><v>1</v></c><c/><x:c r="C1" q:r="qualified"><x:f>1+1</x:f><x:v>2</x:v></x:c></row><row r="2"><c r="B2"><v>3</v></c></row></sheetData></worksheet>"#
        )
        .into_bytes(),
    )
}

fn cell_plan(address: &str, value: i32) -> Plan {
    let address = Address::from_a1(address).expect("test address");
    Plan::cells(BTreeMap::from([(
        address,
        Action::set(Content::from(value)),
    )]))
}

fn assert_store_equivalent(expected: &Store, actual: &Store) {
    assert_eq!(expected.extents(), actual.extents());
    assert_eq!(expected.entries().len(), actual.entries().len());
    for (expected, actual) in expected.entries().iter().zip(actual.entries()) {
        assert_eq!(expected.address, actual.address);
        assert_eq!(expected.cell, actual.cell);
        assert_eq!(expected.style, actual.style);
        assert_eq!(expected.shared_string, actual.shared_string);
        assert_eq!(expected.inline_rich, actual.inline_rich);
        assert_eq!(expected.formula_range, actual.formula_range);
        assert_eq!(expected.shared_formula, actual.shared_formula);
        assert_eq!(expected.cell_metadata, actual.cell_metadata);
        assert_eq!(expected.value_metadata, actual.value_metadata);
    }
    assert_eq!(expected.row_entries(), actual.row_entries());
    assert_eq!(expected.column_entries(), actual.column_entries());
    assert_eq!(expected.defaults(), actual.defaults());
    assert_eq!(expected.merge_ranges(), actual.merge_ranges());
}

#[test]
fn fused_source_pass_matches_eager_store_and_reuses_lossless_layout() {
    let source = source_fixture();
    let eager = parse(source.as_slice(), || Ok(None)).expect("eager worksheet parse");
    let (fused, mut deferred) =
        parse_with_layout(Arc::clone(&source), || Ok(None)).expect("fused worksheet parse");

    assert_store_equivalent(&eager, &fused);
    assert!(deferred.snapshot_error.is_none());
    let layout = deferred.layout.as_ref().expect("ordinary source layout");
    assert_eq!(layout.sheet_data.rows.len(), 2);
    assert_eq!(layout.sheet_data.rows[0].cells.len(), 3);
    assert_eq!(
        layout.sheet_data.rows[0].cells[1].address.a1(),
        "B1",
        "the missing cell reference must use the inferred column"
    );
    assert!(
        layout.sheet_data.rows[0].cells[1].tag.is_none(),
        "the common plain cell must retain the Option<Tag> niche"
    );
    assert!(layout.sheet_data.rows[0].cells[0]
        .tag
        .as_ref()
        .expect("unknown cell attribute tag")
        .attributes
        .iter()
        .any(|attribute| {
            attribute.name.as_ref() == "q:opaque" && attribute.value.as_ref() == "yes"
        }));
    assert!(layout.sheet_data.rows[0].cells[2]
        .tag
        .as_ref()
        .expect("qualified cell tag")
        .attributes
        .iter()
        .any(|attribute| {
            attribute.name.as_ref() == "q:r" && attribute.value.as_ref() == "qualified"
        }));

    let expected = super::super::edit::rewrite(source.as_slice(), "Sheet1", cell_plan("A1", 9))
        .expect("unfused worksheet rewrite");
    let actual = deferred
        .rewrite(source.as_slice(), "Sheet1", cell_plan("A1", 9))
        .expect("layout-backed worksheet rewrite");
    assert_eq!(actual, expected, "fused and unfused output diverged");
    let actual = std::str::from_utf8(&actual).expect("worksheet output UTF-8");
    assert!(actual.contains(
        r#"q:opaque="yes""));
    assert!(actual.contains(r#"q:r="qualified""));
    assert!(actual.contains(r#"<c r="A1" q:opaque="yes"><v>9</v></c>"#
    ));
}

#[test]
fn fused_source_pass_matches_rebound_namespace_cdata_reference_and_shared_formula_state() {
    let source = Arc::new(
        format!(
            r#"<x:worksheet xmlns:x="{S}" xmlns:q="urn:litchi:future" q:root="keep"><x:sheetData><x:row r="1"><x:c r="A1"><x:f t="shared" ref="A1:A2" si="7">B1+$C$1</x:f><x:v><![CDATA[1]]></x:v></x:c><x:c r="A2"><x:f t="shared" si="7"/><x:v>2</x:v></x:c><x:c r="C1"><x:v>&#x33;<![CDATA[.5]]></x:v></x:c><x:c xmlns:x="urn:not-spreadsheet" r="D1"><x:v>ignored</x:v></x:c><x:c r="D1"><x:v>4</x:v></x:c></x:row></x:sheetData></x:worksheet>"#
        )
        .into_bytes(),
    );
    let eager = parse(source.as_slice(), || Ok(None)).expect("eager worksheet parse");
    let (fused, mut deferred) =
        parse_with_layout(Arc::clone(&source), || Ok(None)).expect("fused worksheet parse");

    assert_store_equivalent(&eager, &fused);
    assert!(matches!(
        fused.get(Address::from_a1("C1").expect("address")),
        Some(Cell::Value(crate::cell::Value::Number(value))) if value.as_str() == "3.5"
    ));
    let Some(Cell::Formula(formula)) = fused.get(Address::from_a1("A2").expect("address")) else {
        panic!("expected expanded shared formula")
    };
    assert_eq!(formula.text(), "B2+$C$1");

    let layout = deferred.layout.as_ref().expect("ordinary source layout");
    assert_eq!(
        layout.sheet_data.rows[0]
            .cells
            .iter()
            .map(|cell| cell.address.a1())
            .collect::<Vec<_>>(),
        ["A1", "A2", "C1", "D1"]
    );

    let expected = super::super::edit::rewrite(source.as_slice(), "Sheet1", cell_plan("C1", 8))
        .expect("unfused worksheet rewrite");
    let actual = deferred
        .rewrite(source.as_slice(), "Sheet1", cell_plan("C1", 8))
        .expect("layout-backed worksheet rewrite");
    assert_eq!(actual, expected, "rebound source output diverged");
}

#[test]
fn deferred_layout_falls_back_for_a_different_source_allocation() {
    let source = source_fixture();
    let (_, mut deferred) =
        parse_with_layout(Arc::clone(&source), || Ok(None)).expect("fused worksheet parse");
    let other = source.as_slice().to_vec();
    assert_ne!(source.as_slice().as_ptr(), other.as_ptr());
    let expected = super::super::edit::rewrite(&other, "Sheet1", cell_plan("A1", 9))
        .expect("unfused worksheet rewrite");
    let actual = deferred
        .rewrite(&other, "Sheet1", cell_plan("A1", 9))
        .expect("source identity fallback rewrite");
    assert_eq!(actual, expected);
}

#[test]
fn fused_source_pass_defers_snapshot_only_attribute_failure() {
    let source = Arc::new(
        format!(
        r#"<worksheet xmlns="{S}" xmlns:q="urn:litchi:future"><sheetData><row r="1"><c r="A1" q:opaque="&missing;"><v>1</v></c></row></sheetData></worksheet>"#
        )
        .into_bytes(),
    );
    let eager = parse(source.as_slice(), || Ok(None)).expect("semantic parser ignores q value");
    let (fused, mut deferred) = parse_with_layout(Arc::clone(&source), || Ok(None))
        .expect("semantic parse must complete before snapshot failure");

    assert_store_equivalent(&eager, &fused);
    assert!(
        deferred.layout.is_none(),
        "a failed layout must never be published"
    );
    let error = deferred.snapshot_error.expect("deferred snapshot error");
    assert!(
        error.to_string().contains("missing"),
        "unexpected error: {error}"
    );
    let rewrite_error = deferred
        .rewrite(source.as_slice(), "Sheet1", cell_plan("A1", 2))
        .expect_err("the deferred error must surface at the rewrite boundary");
    assert_eq!(rewrite_error.to_string(), error.to_string());
}

#[test]
fn fused_source_pass_keeps_semantic_error_ahead_of_snapshot_error() {
    let source = format!(
        r#"<worksheet xmlns="{S}" xmlns:q="urn:litchi:future"><sheetData><row r="1"><c r="A1" s="not-a-style" q:opaque="&missing;"/></row></sheetData></worksheet>"#
    );
    let eager = parse(source.as_bytes(), || Ok(None)).expect_err("invalid style");
    let fused = parse_with_layout(Arc::new(source.into_bytes()), || Ok(None))
        .expect_err("semantic error must win over deferred snapshot error");
    assert_eq!(fused.to_string(), eager.to_string());
    assert!(fused.to_string().contains("invalid worksheet cell style"));
    assert!(!fused.to_string().contains("missing"));
}

#[test]
fn fused_source_pass_declines_transformed_alternate_content() {
    let source = format!(
        r#"<worksheet xmlns="{S}" xmlns:mc="{MC}" xmlns:future="urn:litchi:future"><mc:AlternateContent><mc:Choice Requires="future"><sheetData><row r="1"><c r="A1"><v>9</v></c></row></sheetData></mc:Choice><mc:Fallback><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></mc:Fallback></mc:AlternateContent></worksheet>"#
    );
    let source = Arc::new(source.into_bytes());
    let eager = parse(source.as_slice(), || Ok(None)).expect("eager fallback parse");
    let (fused, deferred) =
        parse_with_layout(Arc::clone(&source), || Ok(None)).expect("fallback worksheet parse");

    assert_store_equivalent(&eager, &fused);
    assert!(
        deferred.layout.is_none(),
        "MCE output has no source-byte layout"
    );
    assert!(deferred.snapshot_error.is_none());
    assert!(matches!(
        fused.get(Address::from_a1("A1").expect("address")),
        Some(Cell::Value(crate::cell::Value::Number(value))) if value.as_str() == "7"
    ));
}

#[test]
fn fused_source_pass_preserves_x14ac_preflight_error_precedence() {
    let source = format!(
        r#"<worksheet xmlns="{S}" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:q="urn:litchi:future"><sheetData><row r="1" x14ac:dyDescent="NaN"><c r="A1" q:opaque="&missing;"/></row></sheetData></worksheet>"#
    );
    let error = parse_with_layout(Arc::new(source.into_bytes()), || Ok(None))
        .expect_err("x14ac capture must precede semantic and snapshot work");
    assert!(error
        .to_string()
        .contains("invalid worksheet extension XML"));
    assert!(!error.to_string().contains("missing"));
}

#[test]
fn fused_source_pass_uses_the_injected_snapshot_event_boundary() {
    let source = format!(r#"<worksheet xmlns="{S}"><sheetData/></worksheet>"#);
    let exact = super::super::model::Parser::parse_with_layout_with_event_limit(
        &source,
        || Ok(None),
        x14ac::Values::default(),
        4,
    )
    .expect("four events including Eof are within the limit");
    assert!(exact.1.is_some());
    assert!(exact.2.is_none());

    let mut over = source.trim_end_matches("</worksheet>").to_owned();
    over.push_str("<future/></worksheet>");
    let over = super::super::model::Parser::parse_with_layout_with_event_limit(
        &over,
        || Ok(None),
        x14ac::Values::default(),
        4,
    )
    .expect("semantic parser should continue after snapshot limit");
    assert!(over.0.entries().is_empty());
    assert!(over.1.is_none());
    assert!(over
        .2
        .expect("deferred event-limit error")
        .to_string()
        .contains("worksheet XML exceeds event limit"));

    let over_then_semantic_error = format!(
        r#"<worksheet xmlns="{S}"><sheetData><future/><row r="1"><c r="A1" s="not-a-style"/></row></sheetData></worksheet>"#
    );
    let error = super::super::model::Parser::parse_with_layout_with_event_limit(
        &over_then_semantic_error,
        || Ok(None),
        x14ac::Values::default(),
        4,
    )
    .expect_err("semantic validation must continue after snapshot limit");
    assert!(error.to_string().contains("invalid worksheet cell style"));
    assert!(!error
        .to_string()
        .contains("worksheet XML exceeds event limit"));
}
