//! Differential coverage for a non-empty explicit dimension element.
//!
//! Candidate inclusion target: copy this file to
//! `crates/litchi-xlsx/src/raw/worksheet/edit/package/compact_dimension_expansion.rs`
//! in the isolated candidate (and, when validating the candidate in the live
//! crate, use the same repo-relative destination), then add this ordinary
//! child declaration to `raw/worksheet/edit/package.rs`:
//!
//! ```ignore
//! #[cfg(test)]
//! mod compact_dimension_expansion;
//! ```
//!
//! The test intentionally calls the private compact seam so a fallback cannot
//! make a broken compact dimension writer look correct.

use std::collections::BTreeMap;

use litchi_sheet::Cell as Address;

use super::{rewrite_value_only_with_provenance, try_compact_value_rewrite};
use crate::cell::{Content, Number, Value};
use crate::raw::worksheet;
use crate::raw::worksheet::edit::{Action, collect_compact_layout};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn address(reference: &str) -> Address {
    Address::from_a1(reference).expect("valid worksheet address")
}

fn number_action(value: &str) -> Action {
    Action::set(Content::from(Value::Number(
        Number::new(value).expect("valid worksheet number"),
    )))
}

fn scalar_actions() -> BTreeMap<Address, Action> {
    BTreeMap::from([(address("C1"), number_action("9"))])
}

#[test]
fn compact_dimension_expansion_preserves_explicit_dimension_close() {
    // The source cell is outside the declared one-cell extent.  The edit must
    // expand the opening tag while retaining the explicit closing tag byte for
    // byte.  A compact proof that accidentally spans the whole element loses
    // that close when its writer emits only the replacement opening tag.
    let source = format!(
        r#"<worksheet xmlns="{SML}"><dimension ref="A1"></dimension><sheetData><row r="1"><c r="C1"><v>1</v></c></row></sheetData></worksheet>"#,
    )
    .into_bytes();
    let store = worksheet::parse(&source, || Ok(None)).expect("source worksheet parse");
    assert_eq!(store.entries().len(), 1);
    assert_eq!(store.entries()[0].address, address("C1"));
    let actions = scalar_actions();

    let complete = rewrite_value_only_with_provenance(&source, "Sheet1", scalar_actions())
        .expect("complete worksheet rewrite");
    let proof = collect_compact_layout(&source, store.entries(), &actions)
        .expect("explicit dimension source should admit compact proof");
    let compact = try_compact_value_rewrite(&source, &proof, store.entries(), &actions)
        .expect("eligible scalar edit should use compact writer");

    let expected = format!(
        r#"<worksheet xmlns="{SML}"><dimension ref="A1:C1"></dimension><sheetData><row r="1"><c r="C1"><v>9</v></c></row></sheetData></worksheet>"#,
    )
    .into_bytes();
    assert_eq!(complete.bytes, expected);
    assert_eq!(compact.bytes, expected);
    assert_eq!(compact.bytes, complete.bytes);
    assert_eq!(compact.omitted, complete.omitted);

    let rewritten =
        worksheet::parse(&compact.bytes, || Ok(None)).expect("compact output worksheet parse");
    assert_eq!(rewritten.entries().len(), 1);
    assert_eq!(rewritten.entries()[0].address, address("C1"));
}
