//! Differential coverage for worksheet-child ordering and column envelopes.
//!
//! Candidate inclusion target: copy this file to
//! `crates/litchi-xlsx/src/raw/worksheet/edit/package/compact_layout_order_fallback.rs`
//! in the isolated candidate (and, when validating the candidate in the live
//! crate, use the same repo-relative destination), then add this ordinary
//! child declaration to `raw/worksheet/edit/package.rs`:
//!
//! ```ignore
//! #[cfg(test)]
//! mod compact_layout_order_fallback;
//! ```
//!
//! The complete scanner remains the diagnostic authority for every refusal.

use std::collections::BTreeMap;

use litchi_sheet::Cell as Address;

use super::super::package::{
    rewrite_value_only_with_compact_proof, rewrite_value_only_with_provenance,
    try_compact_value_rewrite,
};
use crate::cell::{Content, Number, Store, Value};
use crate::error::Error;
use crate::raw::worksheet;
use crate::raw::worksheet::edit::{Action, collect_compact_layout};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const CELL_DATA: &str = r#"<sheetData><row r="1"><c r="C1"><v>1</v></c></row></sheetData>"#;

fn address(reference: &str) -> Address {
    Address::from_a1(reference).expect("valid worksheet address")
}

fn worksheet(body: &str) -> Vec<u8> {
    format!(r#"<worksheet xmlns="{SML}">{body}</worksheet>"#).into_bytes()
}

fn number_action(value: &str) -> Action {
    Action::set(Content::from(Value::Number(
        Number::new(value).expect("valid worksheet number"),
    )))
}

fn assert_invalid<T>(result: crate::error::Result<T>, expected: &str, label: &str) {
    match result {
        Err(Error::Invalid(actual)) => assert_eq!(actual, expected, "{label}"),
        Err(_) => panic!("{label}: expected Invalid error"),
        Ok(_) => panic!("{label}: expected complete scanner refusal"),
    }
}

fn authoritative_store(source: &[u8]) -> Store {
    worksheet::parse(source, || Ok(None)).expect("authoritative worksheet Store")
}

#[test]
fn compact_accepts_dimension_defaults_cols_sheet_data_order() {
    // The complete scanner's order is dimension, sheetFormatPr, non-empty
    // cols, then sheetData.  The columns are deliberately retained even
    // though this scalar edit changes only C1.
    let source = worksheet(&format!(
        r#"<dimension ref="A1"/><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1"/></cols>{CELL_DATA}"#,
    ));
    let store = authoritative_store(&source);
    assert_eq!(store.entries().len(), 1);
    assert_eq!(store.entries()[0].address, address("C1"));
    let actions = BTreeMap::from([(address("C1"), number_action("9"))]);

    let complete = rewrite_value_only_with_provenance(&source, "Sheet1", actions.clone())
        .expect("complete ordered worksheet rewrite");
    let proof = collect_compact_layout(&source, store.entries(), &actions)
        .expect("ordered worksheet should admit compact proof");
    let compact = try_compact_value_rewrite(&source, &proof, store.entries(), &actions)
        .expect("ordered worksheet should use compact writer");
    assert_eq!(compact.bytes, complete.bytes);
    assert_eq!(compact.omitted, complete.omitted);

    let wrapped = rewrite_value_only_with_compact_proof(
        &source,
        "Sheet1",
        actions,
        Some(proof),
        store.entries(),
    )
    .expect("compact proof wrapper");
    assert_eq!(wrapped.bytes, complete.bytes);
    assert_eq!(wrapped.omitted, complete.omitted);

    let rewritten = authoritative_store(&compact.bytes);
    assert_eq!(rewritten.entries().len(), 1);
    assert_eq!(rewritten.entries()[0].address, address("C1"));
}

#[test]
fn compact_layout_order_refusals_preserve_complete_errors() {
    // Keep one authoritative Store from a known-good source.  The refusal
    // fixtures contain the same C1 source record, so this isolates child
    // ordering from semantic materialization while still exercising the real
    // parser Store/address sequence.
    let base = worksheet(CELL_DATA);
    let store = authoritative_store(&base);
    let actions = BTreeMap::from([(address("C1"), number_action("9"))]);

    let cases = [
        (
            "dimension after defaults",
            format!(r#"<sheetFormatPr defaultRowHeight="15"/><dimension ref="A1"/>{CELL_DATA}"#,),
            "worksheet dimension must precede sheetFormatPr during edit",
        ),
        (
            "defaults after cols",
            format!(
                r#"<dimension ref="A1"/><cols><col min="1" max="1"/></cols><sheetFormatPr defaultRowHeight="15"/>{CELL_DATA}"#,
            ),
            "worksheet sheetFormatPr appears after column or cell data during edit",
        ),
        (
            "dimension after sheetData",
            format!(r#"{CELL_DATA}<dimension ref="A1"/>"#,),
            "worksheet dimension must precede sheetData during cell edits",
        ),
        (
            "defaults after sheetData",
            format!(r#"{CELL_DATA}<sheetFormatPr defaultRowHeight="15"/>"#,),
            "worksheet sheetFormatPr appears after column or cell data during edit",
        ),
        (
            "cols after sheetData",
            format!(r#"{CELL_DATA}<cols><col min="1" max="1"/></cols>"#,),
            "worksheet cols appears after sheetData during edit",
        ),
        (
            "empty cols",
            format!(r#"<dimension ref="A1"/><cols></cols>{CELL_DATA}"#,),
            "worksheet cols contains no col during edit",
        ),
    ];

    for (label, body, expected) in cases {
        let source = worksheet(&body);
        let proof = collect_compact_layout(&source, store.entries(), &actions);
        assert!(
            proof.is_none(),
            "{label}: malformed order must refuse compact route"
        );

        let complete = rewrite_value_only_with_provenance(&source, "Sheet1", actions.clone());
        assert_invalid(complete, expected, label);
        let fallback = rewrite_value_only_with_compact_proof(
            &source,
            "Sheet1",
            actions.clone(),
            proof,
            store.entries(),
        );
        assert_invalid(fallback, expected, label);
    }
}
