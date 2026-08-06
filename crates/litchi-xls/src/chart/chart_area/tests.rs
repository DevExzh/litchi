use super::super::codec::serialize_chart;
use super::super::package::{build_workbook_fixture, ranges};
use super::super::wire::{CHART, EOF, record};
use super::super::{Chart, Limits, Selector};
use super::*;
use crate::Error;

fn source(rect: Rect) -> Vec<u8> {
    let mut chart = Chart::default();
    chart.x = rect.x;
    chart.y = rect.y;
    chart.width = rect.width;
    chart.height = rect.height;
    serialize_chart(&chart, Limits::default()).expect("valid chart fixture")
}

fn with_unknown(mut input: Vec<u8>) -> Vec<u8> {
    let eof = ranges(&input)
        .expect("framed chart")
        .into_iter()
        .find(|value| value.kind == EOF)
        .expect("EOF")
        .start;
    let opaque = record(0x7777, &[0xA1, 0xB2, 0xC3]).expect("opaque record");
    input.splice(eof..eof, opaque);
    input
}

#[test]
fn fixed_payload_codec_round_trips_wire_values() {
    let value = Snapshot::try_new(Rect {
        x: 0,
        y: 0,
        width: 5 << 16,
        height: 7 << 16,
    })
    .expect("valid chart geometry");
    let payload = encode(&value).expect("fixed payload");
    assert_eq!(payload.len(), 16);
    assert_eq!(decode(&payload).expect("decode").rect(), value.rect());
}

#[test]
fn transaction_is_source_checked_and_reversible() {
    let source = Snapshot::try_new(Rect::default()).expect("valid source");
    let target = Rect {
        x: 0,
        y: 0,
        width: 5_000 << 16,
        height: 2_500 << 16,
    };
    let mut edit = source.edit().expect("transaction");
    edit.set_rect(target).expect("stage");
    let commit = edit.commit().expect("commit");
    let patch = commit.patch();
    let change = patch.change().expect("effective change");
    assert_eq!(change.before(), source.rect());
    assert_eq!(change.after(), target);
    assert_eq!(
        patch.apply(source).expect("apply").snapshot().rect(),
        target
    );
    assert_eq!(
        patch
            .inverse()
            .apply(Snapshot::try_new(target).expect("target"))
            .expect("inverse")
            .snapshot()
            .rect(),
        source.rect()
    );
    assert!(matches!(
        patch.apply(Snapshot::try_new(target).expect("diverged source")),
        Err(Error::UnsafeEdit(_))
    ));
}

#[test]
fn invalid_origin_and_size_are_rejected_before_bytes_change() {
    for rect in [
        Rect {
            x: 1,
            y: 0,
            width: 1,
            height: 1,
        },
        Rect {
            x: 0,
            y: 1,
            width: 1,
            height: 1,
        },
        Rect {
            x: 0,
            y: 0,
            width: -1,
            height: 1,
        },
        Rect {
            x: 0,
            y: 0,
            width: 1,
            height: -1,
        },
    ] {
        assert!(Snapshot::try_new(rect).is_err());
    }
    let source = Snapshot::from_wire(Rect {
        x: -1,
        y: 0,
        width: 1,
        height: 1,
    });
    assert!(source.edit().is_err());
}

#[test]
fn patch_changes_only_the_chart_payload_and_preserves_unknown_offsets() {
    let original = with_unknown(source(Rect::default()));
    let source_snapshot = Snapshot::try_new(Rect::default()).expect("source");
    let target = Rect {
        x: 0,
        y: 0,
        width: 4_500 << 16,
        height: 3_200 << 16,
    };
    let mut edit = source_snapshot.edit().expect("transaction");
    edit.set_rect(target).expect("stage");
    let patch = edit.commit().expect("commit").patch();
    let changed = super::patch(
        &original,
        patch.change().expect("change"),
        Limits::default(),
    )
    .expect("exact patch");

    assert_eq!(changed.len(), original.len());
    let before = ranges(&original).expect("before records");
    let after = ranges(&changed).expect("after records");
    assert_eq!(before.len(), after.len());
    for (before, after) in before.iter().zip(&after) {
        assert_eq!(before.kind, after.kind);
        assert_eq!(before.start, after.start);
        assert_eq!(before.end, after.end);
        let before_bytes = &original[before.start..before.end];
        let after_bytes = &changed[after.start..after.end];
        if before.kind == CHART {
            assert_ne!(before_bytes, after_bytes);
            assert_eq!(
                decode(&changed[after.body_start..after.body_end])
                    .expect("changed chart")
                    .rect(),
                target
            );
        } else {
            assert_eq!(before_bytes, after_bytes);
        }
    }
    assert!(changed.windows(7).any(|window| {
        window[0] == 0x77 && window[1] == 0x77 && window[4..7] == [0xA1, 0xB2, 0xC3]
    }));
}

#[test]
fn package_editor_exposes_contextual_area_edit_without_host_resize() {
    let limits = Limits::default();
    let bytes = build_workbook_fixture(Chart::default(), limits).expect("fixture workbook");
    let mut editor = super::super::Editor::open(bytes, limits).expect("open workbook");
    let target = Rect {
        x: 0,
        y: 0,
        width: 4_800 << 16,
        height: 2_800 << 16,
    };
    editor
        .set_chart_area(
            Selector::Embedded {
                sheet: "Sheet1",
                index: 0,
            },
            target,
        )
        .expect("embedded chart-area payload edit");
    assert_eq!(
        editor
            .get(Selector::Embedded {
                sheet: "Sheet1",
                index: 0,
            })
            .expect("lookup")
            .expect("chart")
            .chart_area()
            .rect(),
        target
    );
}
