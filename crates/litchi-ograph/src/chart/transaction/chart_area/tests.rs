use super::super::super::{Cache, Chart, Context, Rect, RowCol, Series, Value};
use crate::Error;
use crate::chart::cache;
use litchi_biff::{Encoder, Kind, Records};

const UNKNOWN: Kind = Kind::from_wire(0x7777);
const CHART: Kind = Kind::from_wire(0x1002);
const EOF: Kind = Kind::from_wire(0x000A);

fn fixture(rect: Rect) -> crate::chart::Stream {
    let context = Context::graph();
    let mut chart = Chart::new(context).expect("chart");
    chart
        .add_series(Series::new(context))
        .expect("regular series");
    chart
        .add_cache(Cache::graph(
            RowCol::ZERO,
            RowCol::ZERO,
            cache::Ifmt::new(4),
            Value::Number(1.25),
        ))
        .expect("cache");
    chart.set_rect(rect);
    chart.authoring_proven = true;
    chart.encode().expect("authoring fixture")
}

fn with_unknown(source: &crate::chart::Stream) -> crate::chart::Stream {
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
    crate::chart::Stream::open(output.finish()).expect("framed fixture")
}

fn encoded_records(bytes: &[u8]) -> Vec<Vec<u8>> {
    Records::new(bytes)
        .map(|record| record.expect("encoded record").encoded().to_vec())
        .collect()
}

#[test]
fn edits_only_the_fixed_chart_area_and_replays_unknown_records() {
    let source = with_unknown(&fixture(Rect::default()));
    let original = source.as_bytes().to_vec();
    let pointer = source.as_bytes().as_ptr();
    let original_records = encoded_records(&original);
    let chart = Chart::open(source, Context::graph()).expect("parsed chart");
    let target = Rect {
        x: 0,
        y: 0,
        width: 5_000 << 16,
        height: 2_500 << 16,
    };

    let mut edit = chart.edit().expect("transaction");
    edit.set_rect(target).expect("stage chart-area replacement");
    assert_eq!(edit.len(), 1);
    assert!(!edit.is_empty());
    let commit = edit.commit().expect("commit");
    assert_eq!(commit.patch().len(), 1);
    let change = commit.patch().chart_area().expect("chart-area change");
    assert_eq!(change.before(), Rect::default());
    assert_eq!(change.after(), target);

    let changed = commit.into_chart().encode().expect("round trip");
    assert_eq!(changed.as_bytes().as_ptr(), pointer);
    let changed_records = encoded_records(changed.as_bytes());
    assert_eq!(changed_records.len(), original_records.len());
    for (before, after) in original_records.iter().zip(changed_records.iter()) {
        if before.get(0..2) == Some(&CHART.get().to_le_bytes()) {
            assert_ne!(before, after);
        } else {
            assert_eq!(before, after);
        }
    }
    let reparsed = Chart::open(changed, Context::graph()).expect("reparse");
    assert_eq!(reparsed.rect(), target);
    assert_eq!(reparsed.unknown().len(), 1);
    assert_eq!(reparsed.unknown()[0].kind(), UNKNOWN);
}

#[test]
fn inverse_restores_the_exact_source_allocation() {
    let source = with_unknown(&fixture(Rect::default()));
    let original = source.as_bytes().to_vec();
    let chart = Chart::open(source, Context::graph()).expect("parsed chart");
    let target = Rect {
        x: 0,
        y: 0,
        width: 4_500 << 16,
        height: 3_200 << 16,
    };
    let mut edit = chart.edit().expect("transaction");
    edit.set_rect(target).expect("stage");
    let commit = edit.commit().expect("commit");
    let restored = commit
        .patch()
        .inverse()
        .apply(commit.into_chart())
        .expect("inverse commit")
        .into_chart()
        .encode()
        .expect("restored stream");
    assert_eq!(restored.as_bytes(), original.as_slice());
}

#[test]
fn invalid_or_conflicting_area_edits_fail_before_publication() {
    let source = with_unknown(&fixture(Rect::default()));
    let original = source.as_bytes().to_vec();
    let chart = Chart::open(source, Context::graph()).expect("parsed chart");
    let mut edit = chart.edit().expect("transaction");
    assert!(matches!(
        edit.set_rect(Rect {
            x: 1,
            y: 0,
            width: 1,
            height: 1,
        }),
        Err(Error::InvalidModel {
            field: "chart area",
            reason: "Chart x and y must be zero",
        })
    ));
    assert!(edit.is_empty());
    assert_eq!(
        edit.commit()
            .expect("rejected request leaves a no-op transaction")
            .into_chart()
            .encode()
            .expect("exact replay")
            .as_bytes(),
        original.as_slice()
    );

    let source = fixture(Rect::default());
    let chart = Chart::open(source, Context::graph()).expect("source chart");
    let target = Rect {
        x: 0,
        y: 0,
        width: 4_200 << 16,
        height: 3_000 << 16,
    };
    let mut edit = chart.edit().expect("transaction");
    edit.set_rect(target).expect("stage");
    let patch = edit.commit().expect("commit").patch().clone();
    let other = Chart::open(
        fixture(Rect {
            x: 0,
            y: 0,
            width: 4_100 << 16,
            height: 3_000 << 16,
        }),
        Context::graph(),
    )
    .expect("different source chart");
    assert!(matches!(
        patch.apply(other),
        Err(Error::UnsupportedMutation {
            operation: "chart-area-patch",
            reason: "patch source rectangle does not match the target snapshot",
        })
    ));
}
