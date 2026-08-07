#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic by design"
)]

use super::super::{Cache, Chart, Context, RowCol, Series, Value, ValueRef, cache};
use super::{CacheValue, Error, Identity};
use crate::chart::Stream;
use litchi_biff::{Encoder, Kind, Records};

const UNKNOWN: Kind = Kind::from_wire(0x7777);
const EOF: Kind = Kind::from_wire(0x000A);
const GRAPH_NUMBER: Kind = Kind::from_wire(0x0003);

fn fixture(row: u16, col: u16, value: Value) -> Stream {
    let context = Context::graph();
    let mut chart = Chart::new(context).expect("chart");
    chart
        .add_series(Series::new(context))
        .expect("regular series");
    chart
        .add_cache(Cache::graph(
            RowCol::new(row).expect("row"),
            RowCol::new(col).expect("column"),
            cache::Ifmt::new(4),
            value,
        ))
        .expect("cache");
    chart.authoring_proven = true;
    let source = chart.encode().expect("authoring fixture");

    let mut output = Encoder::new();
    for item in Records::new(source.as_bytes()) {
        let record = item.expect("fixture record");
        if record.kind() == EOF {
            output
                .push(UNKNOWN, &[0xA1, 0xB2, 0xC3])
                .expect("opaque record");
        }
        output.push_ref(record).expect("record replay");
    }
    Stream::open(output.finish()).expect("framed fixture")
}

fn encoded_records(bytes: &[u8]) -> Vec<Vec<u8>> {
    Records::new(bytes)
        .map(|record| record.expect("encoded record").encoded().to_vec())
        .collect()
}

#[test]
fn graph_patch_preserves_identity_unknown_order_and_source_allocation() {
    let source = fixture(2, 7, Value::Number(3.5));
    let pointer = source.as_bytes().as_ptr();
    let original = source.as_bytes().to_vec();
    let original_records = encoded_records(&original);

    let chart = Chart::open(source, Context::graph()).expect("parsed chart");
    let mut edit = chart.edit().expect("transaction");
    edit.set_cache_value(0, CacheValue::Graph(Value::Number(7.25)))
        .expect("stage numeric replacement");
    let commit = edit.commit().expect("commit");
    assert_eq!(commit.patch().len(), 1);
    assert!(matches!(
        commit.patch().changes()[0].identity(),
        Identity::Graph { row, col, ifmt }
            if row.get() == 2 && col.get() == 7 && ifmt.get() == 4
    ));
    assert_eq!(commit.chart().caches()[0].value(), ValueRef::Number(7.25));

    let changed = commit.into_chart().encode().expect("round trip");
    assert_eq!(changed.as_bytes().as_ptr(), pointer);
    let changed_records = encoded_records(changed.as_bytes());
    assert_eq!(changed_records.len(), original_records.len());
    for (before, after) in original_records.iter().zip(changed_records.iter()) {
        if before.get(0..2) == Some(&GRAPH_NUMBER.get().to_le_bytes()) {
            assert_ne!(before, after);
        } else {
            assert_eq!(before, after);
        }
    }

    let reparsed = Chart::open(changed, Context::graph()).expect("reparse");
    assert_eq!(reparsed.caches()[0].value(), ValueRef::Number(7.25));
    assert_eq!(reparsed.unknown().len(), 1);
    assert_eq!(reparsed.unknown()[0].kind(), UNKNOWN);
}

#[test]
fn inverse_patch_restores_the_exact_graph_stream() {
    let source = fixture(1, 3, Value::Number(11.0));
    let original = source.as_bytes().to_vec();
    let chart = Chart::open(source, Context::graph()).expect("parsed chart");
    let mut edit = chart.edit().expect("transaction");
    edit.set_cache_value(0, CacheValue::Graph(Value::Number(12.0)))
        .expect("stage");
    let commit = edit.commit().expect("commit");
    let inverse = commit.patch().inverse();
    let restored = inverse
        .apply(commit.into_chart())
        .expect("inverse commit")
        .into_chart()
        .encode()
        .expect("restored stream");
    assert_eq!(restored.as_bytes(), original.as_slice());
}

#[test]
fn no_op_and_unsafe_value_replacements_are_rejected_without_reencoding() {
    let source = fixture(0, 0, Value::Number(4.0));
    let pointer = source.as_bytes().as_ptr();
    let original = source.as_bytes().to_vec();
    let chart = Chart::open(source, Context::graph()).expect("parsed chart");
    let mut edit = chart.edit().expect("transaction");
    edit.set_cache_value(0, CacheValue::Graph(Value::Number(4.0)))
        .expect("stage no-op");
    let commit = edit.commit().expect("no-op commit");
    assert!(commit.patch().is_empty());
    let replay = commit.into_chart().encode().expect("no-op replay");
    assert_eq!(replay.as_bytes(), original.as_slice());
    assert_eq!(replay.as_bytes().as_ptr(), pointer);

    let text_chart =
        Chart::open(fixture(0, 0, Value::Text("a".into())), Context::graph()).expect("text chart");
    let mut text = text_chart.edit().expect("text transaction");
    text.set_cache_value(0, CacheValue::Graph(Value::Text("long".into())))
        .expect("stage text replacement");
    assert!(matches!(
        text.commit(),
        Err(Error::UnsupportedMutation {
            operation: "cache-value-patch",
            reason: "replacement changes the physical cache record length or class"
        })
    ));
}

#[test]
fn typed_xnum_and_identity_guards_fail_before_source_mutation() {
    let nan = Chart::open(fixture(0, 0, Value::Number(1.0)), Context::graph())
        .expect("numeric chart")
        .edit()
        .expect("transaction");
    let mut nan = nan;
    assert!(matches!(
        nan.set_cache_value(0, CacheValue::Graph(Value::Number(f64::NAN))),
        Err(Error::InvalidModel {
            field: "cache value",
            ..
        })
    ));

    let first_chart =
        Chart::open(fixture(0, 0, Value::Number(1.0)), Context::graph()).expect("first chart");
    let mut first_edit = first_chart.edit().expect("first transaction");
    first_edit
        .set_cache_value(0, CacheValue::Graph(Value::Number(2.0)))
        .expect("stage");
    let first = first_edit.commit().expect("first commit");
    let patch = first.patch().clone();
    let other =
        Chart::open(fixture(0, 1, Value::Number(1.0)), Context::graph()).expect("different chart");
    assert!(matches!(
        patch.apply(other),
        Err(Error::UnsupportedMutation {
            operation: "cache-value-patch",
            reason: "patch cache identity does not match the target snapshot"
        })
    ));
}
