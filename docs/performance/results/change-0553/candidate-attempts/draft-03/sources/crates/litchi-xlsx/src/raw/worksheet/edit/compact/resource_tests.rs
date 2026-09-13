//! Private child of edit::compact; copy to compact/resource_tests.rs.
use std::collections::BTreeMap;
use std::mem::size_of;

use super::super::package::try_compact_value_rewrite;
use super::{MAX_PROOF_BYTES, collect_compact_layout, reserve_vec};
use crate::cell::{Content, Number, Value};
use crate::raw::worksheet::edit::Action;
use crate::raw::worksheet::{self, MAX_SHARED_SOURCE_BYTES, shared_event_bound_within_cap};
use litchi_sheet::Cell as Address;

const HEAD: &str =
    r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#;
const TAIL: &str = "</sheetData></worksheet>";

fn source(rows: usize) -> Vec<u8> {
    let mut text = HEAD.to_owned();
    for row in 1..=rows {
        text.push_str(&format!(
            r#"<row r="{row}"><c r="A{row}"><v>1</v></c></row>"#
        ));
    }
    text.push_str(TAIL);
    text.into_bytes()
}

fn actions() -> BTreeMap<Address, Action> {
    BTreeMap::from([(
        Address::from_a1("A1").expect("address"),
        Action::set(Content::from(Value::Number(
            Number::new("9").expect("number"),
        ))),
    )])
}

#[test]
fn compact_reservation_refuses_growth_without_changing_existing_storage() {
    let mut values = Vec::<u64>::new();
    let mut charged = MAX_PROOF_BYTES - 16 * size_of::<u64>();
    assert!(reserve_vec(&mut values, 16, &mut charged));
    assert_eq!(charged, MAX_PROOF_BYTES);
    values.resize(values.capacity(), 7);
    let before = (values.as_ptr(), values.len(), values.capacity(), charged);
    assert!(!reserve_vec(&mut values, 16, &mut charged));
    assert_eq!(
        before,
        (values.as_ptr(), values.len(), values.capacity(), charged)
    );
    assert!(values.iter().all(|value| *value == 7));
}

#[test]
fn compact_reservation_refuses_overflow_before_allocating() {
    for (growth, initial) in [(usize::MAX, 0), (16, usize::MAX)] {
        let mut values = Vec::<u64>::new();
        let mut charged = initial;
        assert!(!reserve_vec(&mut values, growth, &mut charged));
        assert_eq!(values.capacity(), 0);
        assert_eq!(charged, initial);
    }
}

#[test]
fn compact_source_byte_limit_is_inclusive() {
    let mut bytes = source(1);
    let store = worksheet::parse(&bytes, || Ok(None)).expect("source parse");
    let edits = actions();
    bytes.resize(MAX_SHARED_SOURCE_BYTES, b' ');
    assert!(shared_event_bound_within_cap(&bytes));
    assert!(collect_compact_layout(&bytes, store.entries(), &edits).is_some());
    bytes.push(b' ');
    assert!(collect_compact_layout(&bytes, store.entries(), &edits).is_none());
}

#[test]
fn compact_proof_refuses_equal_bytes_from_different_source_or_store_allocations() {
    let bytes = source(1);
    let store = worksheet::parse(&bytes, || Ok(None)).expect("source parse");
    let edits = actions();
    let proof = collect_compact_layout(&bytes, store.entries(), &edits).expect("compact proof");
    assert!(try_compact_value_rewrite(&bytes, &proof, store.entries(), &edits).is_some());
    let copied_source = bytes.clone();
    let copied_entries = store.entries().to_vec();
    assert_ne!(bytes.as_ptr(), copied_source.as_ptr());
    assert_ne!(store.entries().as_ptr(), copied_entries.as_ptr());
    assert!(try_compact_value_rewrite(&copied_source, &proof, store.entries(), &edits).is_none());
    assert!(try_compact_value_rewrite(&bytes, &proof, &copied_entries, &edits).is_none());
    assert!(
        try_compact_value_rewrite(&bytes[..bytes.len() - 1], &proof, store.entries(), &edits)
            .is_none()
    );
    assert!(try_compact_value_rewrite(&bytes, &proof, &[], &edits).is_none());
}

#[test]
fn compact_many_rows_reaches_metadata_refusal_below_source_and_event_limits() {
    let edits = actions();
    for (rows, admitted) in [(4_000, true), (12_000, false)] {
        let bytes = source(rows);
        assert!(bytes.len() < MAX_SHARED_SOURCE_BYTES);
        assert!(shared_event_bound_within_cap(&bytes));
        let store = worksheet::parse(&bytes, || Ok(None)).expect("many-row source parse");
        assert_eq!(store.entries().len(), rows);
        assert_eq!(
            collect_compact_layout(&bytes, store.entries(), &edits).is_some(),
            admitted
        );
    }
}
