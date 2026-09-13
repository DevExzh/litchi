//! Extended differential coverage for the compact source-cell proof.
//!
//! This file is included as a private child of `cell_values::source_proof`.
//! The direct compact seam makes the accepted route observable; the public
//! wrapper is checked separately for the cases that must fall back to the
//! complete scanner and writer.

use std::collections::BTreeMap;

use litchi_sheet::Cell as Address;

use super::super::validation::{worksheet_xml, worksheet_xml_and_parse_source};
use crate::cell::{Content, Number, Text, Value};
use crate::formula::Formula;
use crate::raw::worksheet::edit::{
    Action, Payload, StyleEffect, ValueOnlyRewrite, reduced_readback,
    rewrite_value_only_with_compact_proof, rewrite_value_only_with_provenance,
    try_compact_value_rewrite,
};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn worksheet(body: &str) -> Vec<u8> {
    format!(r#"<worksheet xmlns="{SML}">{body}</worksheet>"#).into_bytes()
}

fn address(reference: &str) -> Address {
    Address::from_a1(reference).expect("valid worksheet address")
}

fn number_action(value: &str) -> Action {
    Action::set(Content::from(Value::Number(
        Number::new(value).expect("valid worksheet number"),
    )))
}

fn formula_action(value: &str) -> Action {
    Action::set(Content::from(
        Formula::new(value).expect("valid worksheet formula"),
    ))
}

fn assert_rewrites_equal(
    compact: ValueOnlyRewrite<'_>,
    complete: ValueOnlyRewrite<'_>,
) {
    worksheet_xml(&compact.bytes).expect("compact emitted worksheet validation");
    assert_eq!(&compact.bytes, &complete.bytes);
    assert_eq!(&compact.omitted, &complete.omitted);

    let compact_reduced =
        reduced_readback(&compact.bytes, &compact.omitted).expect("compact reduced readback");
    let complete_reduced =
        reduced_readback(&complete.bytes, &complete.omitted).expect("complete reduced readback");
    assert_eq!(&compact_reduced, &complete_reduced);

    let compact_store = crate::raw::worksheet::parse(&compact_reduced, || Ok(None))
        .expect("compact reduced parse");
    let complete_store = crate::raw::worksheet::parse(&complete_reduced, || Ok(None))
        .expect("complete reduced parse");
    assert_eq!(
        compact_store.entries().len(),
        complete_store.entries().len()
    );
    for (compact, complete) in compact_store
        .entries()
        .iter()
        .zip(complete_store.entries())
    {
        assert_eq!(compact.address, complete.address);
        assert_eq!(compact.cell, complete.cell);
        assert_eq!(compact.style, complete.style);
    }
}

fn assert_accepted_differential(source: &[u8], actions: BTreeMap<Address, Action>) {
    let (store, proof) = worksheet_xml_and_parse_source(source).expect("source parse");
    let proof = proof.expect("formula-free source admits compact proof");
    let compact = try_compact_value_rewrite(source, &proof, store.entries(), &actions)
        .expect("eligible actions use the compact writer");
    let complete = rewrite_value_only_with_provenance(source, "Sheet1", actions)
        .expect("complete rewrite");
    assert_rewrites_equal(compact, complete);
}

fn assert_fallback_differential(
    source: &[u8],
    actions: BTreeMap<Address, Action>,
    proof: Option<&super::CompactLayout>,
    entries: &[crate::cell::Stored],
) {
    let compact = rewrite_value_only_with_compact_proof(
        source,
        "Sheet1",
        actions.clone(),
        proof,
        entries,
    );
    let complete = rewrite_value_only_with_provenance(source, "Sheet1", actions);
    match (compact, complete) {
        (Ok(compact), Ok(complete)) => assert_rewrites_equal(compact, complete),
        (Err(compact), Err(complete)) => {
            assert_eq!(format!("{compact:?}"), format!("{complete:?}"));
        },
        (compact, complete) => {
            panic!("compact/fallback result mismatch: {compact:?} versus {complete:?}");
        },
    }
}

#[test]
fn compact_set_formula_and_discontiguous_existing_cells_match_complete() {
    let source = worksheet(
        r#"<dimension ref="A1:D4"/><sheetData><row r="1"><c r="A1" t="n"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c><c r="C2" s="2"><v>4</v></c></row><row r="4"><c r="D4"><v>5</v></c></row></sheetData>"#,
    );
    let actions = BTreeMap::from([
        (address("B1"), formula_action("A1+1")),
        (address("C2"), Action::style(7)),
        (address("D4"), number_action("50")),
    ]);

    assert_accepted_differential(&source, actions);
}

#[test]
fn compact_clear_and_clear_if_present_match_for_styled_and_empty_cells() {
    let source = worksheet(
        r#"<sheetData><row r="1"><c r="A1" s="4"><v>1</v></c><c r="B1" s="5"/></row><row r="2"><c r="A2"/><c r="B2"/></row></sheetData>"#,
    );
    let actions = BTreeMap::from([
        (address("A1"), Action::clear(true)),
        (address("B1"), Action::clear(false)),
        (address("A2"), Action::clear(false)),
        (address("B2"), Action::clear(true)),
    ]);

    assert_accepted_differential(&source, actions);
}

#[test]
fn compact_style_set_reset_and_orthogonal_payload_match_complete() {
    let source = worksheet(
        r#"<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>2</v></c><c r="C1" s="3"><v>3</v></c></row></sheetData>"#,
    );
    let actions = BTreeMap::from([
        (address("A1"), Action::style(9)),
        (address("B1"), Action::reset_style()),
        (
            address("C1"),
            Action::Update {
                payload: Some(Payload::Set(Content::from(Value::Number(
                    Number::new("30").expect("valid worksheet number"),
                )))),
                style: Some(StyleEffect::Set(8)),
            },
        ),
    ]);

    assert_accepted_differential(&source, actions);
}

#[test]
fn compact_normalizes_prefixed_changed_cell_attributes_like_complete_writer() {
    let source = format!(
        r#"<x:worksheet xmlns:x="{SML}" xmlns:y="{SML}"><x:sheetData><x:row r="1"><y:c r="A1" s="&#x31;" t="&#x6e;"><y:v><![CDATA[7]]></y:v></y:c><y:c r="C1" s="&#x32;"><y:v>2</y:v></y:c></x:row><x:row r="3"><y:c r="B3"><y:v>3</y:v></y:c></x:row></x:sheetData></x:worksheet>"#,
    )
    .into_bytes();
    let actions = BTreeMap::from([
        (address("A1"), Action::style(4)),
        (address("B3"), number_action("33")),
    ]);

    assert_accepted_differential(&source, actions);
}

#[test]
fn formula_source_and_shared_string_action_use_complete_fallback() {
    let formula_source = worksheet(
        r#"<sheetData><row r="1"><c r="A1"><f>A1+1</f><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>"#,
    );
    let (formula_store, formula_proof) =
        worksheet_xml_and_parse_source(&formula_source).expect("formula source parse");
    assert!(formula_proof.is_none());
    assert_fallback_differential(
        &formula_source,
        BTreeMap::from([(address("B1"), number_action("9"))]),
        formula_proof.as_ref(),
        formula_store.entries(),
    );

    let shared_string_source = worksheet(
        r#"<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>"#,
    );
    let (shared_string_store, shared_string_proof) =
        worksheet_xml_and_parse_source(&shared_string_source).expect("shared string source parse");
    let shared_string_proof = shared_string_proof.expect("formula-free source proof");
    let shared_string_action = Action::Update {
        payload: Some(Payload::SharedString {
            index: 0,
            text: Text::from("shared"),
        }),
        style: None,
    };
    let shared_string_actions =
        BTreeMap::from([(address("A1"), shared_string_action.clone())]);
    assert!(try_compact_value_rewrite(
        &shared_string_source,
        &shared_string_proof,
        shared_string_store.entries(),
        &shared_string_actions,
    )
    .is_none());
    assert_fallback_differential(
        &shared_string_source,
        shared_string_actions,
        Some(&shared_string_proof),
        shared_string_store.entries(),
    );
}
