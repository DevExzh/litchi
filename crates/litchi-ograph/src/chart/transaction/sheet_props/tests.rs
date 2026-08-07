#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic by design"
)]

use super::super::super::{Cache, Chart, Context, Props, RowCol, Series, Value};
use crate::Error;
use crate::chart::Stream;
use litchi_biff::{Encoder, Kind, Records};

const UNKNOWN: Kind = Kind::from_wire(0x7777);
const EOF: Kind = Kind::from_wire(0x000A);
const SHT_PROPS: Kind = Kind::from_wire(0x1044);

fn graph_fixture() -> Stream {
    let context = Context::graph();
    let mut chart = Chart::new(context).expect("chart");
    chart
        .add_series(Series::new(context))
        .expect("regular series");
    chart
        .add_cache(Cache::graph(
            RowCol::ZERO,
            RowCol::ZERO,
            super::super::super::cache::Ifmt::new(4),
            Value::Number(1.25),
        ))
        .expect("cache");
    chart.set_props(Props {
        flags: 1 | (1 << 1) | (1 << 3) | (1 << 4),
        plot_area: true,
    });
    chart.authoring_proven = true;
    chart.encode().expect("Graph fixture")
}

fn excel_fixture() -> Stream {
    let context = Context::excel();
    let mut chart = Chart::new(context).expect("chart");
    chart.set_props(Props {
        flags: 1 | (1 << 1) | (1 << 3) | (1 << 4),
        plot_area: true,
    });
    chart.authoring_proven = true;
    chart.encode().expect("Excel fixture")
}

fn with_unknown(source: &Stream) -> Stream {
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

fn with_prefix_unknown(source: &Stream) -> Stream {
    let mut output = Encoder::new();
    for item in Records::new(source.as_bytes()) {
        let record = item.expect("fixture record");
        if record.kind() == SHT_PROPS {
            output.push(UNKNOWN, &[0xFE, 0xED]).expect("prefix record");
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

fn target(before: Props, blank: u32) -> Props {
    Props {
        flags: (before.flags & !0x00FF_0000) | (blank << 16),
        plot_area: before.plot_area,
    }
}

#[test]
fn graph_patches_only_existing_sht_props_and_preserves_unknown_order() {
    let source = with_unknown(&graph_fixture());
    let original = source.as_bytes().to_vec();
    let pointer = source.as_bytes().as_ptr();
    let original_records = encoded_records(&original);
    let chart = Chart::open(source, Context::graph()).expect("parsed Graph chart");
    let before = chart.props();
    let after = target(before, 2);

    let mut edit = chart.edit().expect("transaction");
    edit.set_props(after).expect("stage ShtProps replacement");
    assert_eq!(edit.len(), 1);
    let commit = edit.commit().expect("commit");
    let change = commit.patch().sheet_props().expect("ShtProps change");
    assert_eq!(change.before(), before);
    assert_eq!(change.after(), after);

    let changed = commit.into_chart().encode().expect("round trip");
    assert_eq!(changed.as_bytes().as_ptr(), pointer);
    let changed_records = encoded_records(changed.as_bytes());
    assert_eq!(changed_records.len(), original_records.len());
    for (before, after) in original_records.iter().zip(changed_records.iter()) {
        if before.get(0..2) == Some(&SHT_PROPS.get().to_le_bytes()) {
            assert_ne!(before, after);
        } else {
            assert_eq!(before, after);
        }
    }
    let reparsed = Chart::open(changed, Context::graph()).expect("reparse");
    assert_eq!(reparsed.props(), after);
    assert_eq!(reparsed.unknown().len(), 1);
    assert_eq!(reparsed.unknown()[0].kind(), UNKNOWN);
}

#[test]
fn excel_patches_sht_props_without_changing_plot_area_topology() {
    let source = with_unknown(&excel_fixture());
    let original = source.as_bytes().to_vec();
    let original_records = encoded_records(&original);
    let chart = Chart::open(source, Context::excel()).expect("parsed Excel chart");
    let before = chart.props();
    let after = target(before, 1);

    let mut edit = chart.edit().expect("transaction");
    edit.set_props(after).expect("stage");
    let commit = edit.commit().expect("commit");
    assert_eq!(commit.patch().len(), 1);
    assert_eq!(commit.chart().props(), after);

    let changed = commit.into_chart().encode().expect("round trip");
    let changed_records = encoded_records(changed.as_bytes());
    assert_eq!(changed_records.len(), original_records.len());
    assert_eq!(
        changed_records
            .iter()
            .filter(|record| record.get(0..2) == Some(&SHT_PROPS.get().to_le_bytes()))
            .count(),
        1
    );
    assert_eq!(
        changed_records
            .iter()
            .filter(|record| record.get(0..2) == Some(&Kind::from_wire(0x1035).get().to_le_bytes()))
            .count(),
        1
    );
    for (before, after) in original_records.iter().zip(changed_records.iter()) {
        if before.get(0..2) == Some(&SHT_PROPS.get().to_le_bytes()) {
            assert_ne!(before, after);
        } else {
            assert_eq!(before, after);
        }
    }
}

#[test]
fn inverse_and_source_identity_guards_restore_or_refuse_exactly() {
    let source = with_unknown(&graph_fixture());
    let original = source.as_bytes().to_vec();
    let chart = Chart::open(source, Context::graph()).expect("parsed chart");
    let after = target(chart.props(), 1);
    let mut edit = chart.edit().expect("transaction");
    edit.set_props(after).expect("stage");
    let commit = edit.commit().expect("commit");
    let inverse = commit.patch().inverse();
    let restored = inverse
        .apply(commit.into_chart())
        .expect("inverse commit")
        .into_chart()
        .encode()
        .expect("restored stream");
    assert_eq!(restored.as_bytes(), original.as_slice());

    let other = Chart::open(with_unknown(&excel_fixture()), Context::excel())
        .expect("different producer chart");
    assert!(matches!(
        inverse.apply(other),
        Err(Error::UnsupportedMutation {
            operation: "sheet-props-patch",
            reason: "patch source ShtProps value does not match the target snapshot",
        })
    ));

    let prefixed = Chart::open(with_prefix_unknown(&graph_fixture()), Context::graph())
        .expect("same semantic chart with a shifted source record");
    let mut prefixed_edit = prefixed.edit().expect("transaction");
    prefixed_edit.set_props(after).expect("stage");
    let prefixed_changed = prefixed_edit.commit().expect("commit").into_chart();
    assert!(matches!(
        inverse.apply(prefixed_changed),
        Err(Error::UnsupportedMutation {
            operation: "sheet-props-patch",
            reason: "source ShtProps record identity does not match the patch",
        })
    ));
}

#[test]
fn invalid_flags_and_plot_area_changes_are_refused_before_publication() {
    let source = graph_fixture();
    let original = source.as_bytes().to_vec();
    let chart = Chart::open(source, Context::graph()).expect("parsed chart");
    let before = chart.props();
    let mut edit = chart.edit().expect("transaction");

    assert!(matches!(
        edit.set_props(Props {
            flags: before.flags | (1 << 2),
            plot_area: before.plot_area,
        }),
        Err(Error::InvalidModel {
            field: "sheet properties",
            ..
        })
    ));
    assert!(matches!(
        edit.set_props(Props {
            flags: before.flags,
            plot_area: !before.plot_area,
        }),
        Err(Error::UnsupportedMutation {
            operation: "sheet-props-patch",
            reason: "PlotArea record presence cannot change in a fixed-record transaction",
        })
    ));
    assert!(edit.is_empty());
    assert_eq!(
        edit.commit()
            .expect("no-op commit")
            .into_chart()
            .encode()
            .expect("exact replay")
            .as_bytes(),
        original.as_slice()
    );
}
