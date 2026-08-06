use super::{Metadata, Snapshot};
use crate::chart::{Cache, Chart, Context, Count, DataKind, RowCol, Series, Stream, Value, cache};
use crate::{Error, Result};
use litchi_biff::{Encoder, Kind, Records};

const UNKNOWN: Kind = Kind::from_wire(0x7777);
const SERIES: Kind = Kind::from_wire(0x1003);

fn fixture(category: DataKind, category_count: u16, unknown_payload: &[u8]) -> Stream {
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
            Value::Number(1.0),
        ))
        .expect("cache");
    chart.series_mut()[0].category_kind = category;
    chart.series_mut()[0].category_count = Count::new(category_count).expect("count");
    chart.series_mut()[0].value_count = Count::new(2).expect("count");
    chart.series_mut()[0].bubble_count = Count::new(1).expect("count");
    chart.authoring_proven = true;
    let source = chart.encode().expect("authoring fixture");
    let mut output = Encoder::new();
    for item in Records::new(source.as_bytes()) {
        let record = item.expect("fixture record");
        if record.kind() == SERIES {
            output
                .push(UNKNOWN, unknown_payload)
                .expect("unknown record");
        }
        output.push_ref(record).expect("record replay");
    }
    Stream::open(output.finish()).expect("framed fixture")
}

fn series_payload(bytes: &[u8]) -> Vec<u8> {
    Records::new(bytes)
        .find_map(|item| {
            let record = item.expect("record");
            (record.kind() == SERIES).then(|| record.payload().to_vec())
        })
        .expect("Series record")
}

#[test]
fn edits_series_metadata_in_place_and_preserves_unknown_records() -> Result<()> {
    let source = fixture(DataKind::Text, 1, &[0xA1, 0xB2, 0xC3]);
    let original = source.as_bytes().to_vec();
    let chart = Chart::open(source, Context::graph())?;
    let snapshot = Snapshot::from_chart(chart)?;
    assert_eq!(snapshot.get(0).expect("metadata").category_count().get(), 1);
    let mut edit = snapshot.edit();
    edit.set(
        0,
        Metadata::new(
            DataKind::Numeric,
            Count::new(7).expect("count"),
            Count::new(8).expect("count"),
            Count::new(2).expect("count"),
        ),
    )?;
    let commit = edit.commit()?;
    assert_eq!(commit.patch().len(), 1);
    assert_eq!(commit.chart().series()[0].category_kind, DataKind::Numeric);
    assert_eq!(commit.chart().series()[0].value_count.get(), 8);
    let changed = commit.into_chart().encode()?;
    assert_eq!(changed.as_bytes().len(), original.len());
    assert_ne!(
        series_payload(&original),
        series_payload(changed.as_bytes())
    );
    let unknowns: Vec<_> = Records::new(changed.as_bytes())
        .filter_map(|item| {
            let record = item.expect("record");
            (record.kind() == UNKNOWN).then(|| record.payload().to_vec())
        })
        .collect();
    assert_eq!(unknowns, vec![vec![0xA1, 0xB2, 0xC3]]);
    Ok(())
}

#[test]
fn no_op_commit_replays_the_exact_source() -> Result<()> {
    let source = fixture(DataKind::Text, 1, &[0x10, 0x20]);
    let pointer = source.as_bytes().as_ptr();
    let original = source.as_bytes().to_vec();
    let chart = Chart::open(source, Context::graph())?;
    let snapshot = Snapshot::from_chart(chart)?;
    let current = snapshot.get(0).expect("metadata");
    let mut edit = snapshot.edit();
    edit.set(0, current)?;
    let commit = edit.commit()?;
    assert!(commit.patch().is_empty());
    let replay = commit.into_chart().encode()?;
    assert_eq!(replay.as_bytes(), original.as_slice());
    assert_eq!(replay.as_bytes().as_ptr(), pointer);
    Ok(())
}

#[test]
fn stale_sources_are_rejected_and_inverse_restores_exact_bytes() -> Result<()> {
    let source = fixture(DataKind::Text, 1, &[0x01]);
    let original = source.as_bytes().to_vec();
    let chart = Chart::open(source, Context::graph())?;
    let mut edit = Snapshot::from_chart(chart)?.edit();
    edit.set(
        0,
        Metadata::new(
            DataKind::Text,
            Count::new(3).expect("count"),
            Count::new(2).expect("count"),
            Count::new(1).expect("count"),
        ),
    )?;
    let commit = edit.commit()?;
    let (changed_chart, patch) = commit.into_parts();
    let changed_stream = changed_chart.encode()?;
    let changed = changed_stream.as_bytes().to_vec();
    let inverse = patch.inverse();
    let changed_target = Chart::open(changed_stream, Context::graph())?;
    let restored = inverse.apply(changed_target)?.into_chart().encode()?;
    assert_eq!(restored.as_bytes(), original.as_slice());

    let stale = Chart::open(fixture(DataKind::Text, 1, &[0x02]), Context::graph())?;
    assert!(matches!(
        patch.apply(stale),
        Err(Error::UnsupportedMutation {
            operation: "series-metadata-patch",
            reason: "patch source fingerprint does not match the target snapshot"
        })
    ));
    assert_ne!(changed, original);
    Ok(())
}

#[test]
fn invalid_metadata_fails_before_publication() -> Result<()> {
    let source = fixture(DataKind::Text, 1, &[0xAA]);
    let chart = Chart::open(source, Context::graph())?;
    let snapshot = Snapshot::from_chart(chart)?;
    let mut edit = snapshot.edit();
    let error = edit
        .set(
            0,
            Metadata::new(
                DataKind::Text,
                Count::new(0x0FA0).expect("Count permits the wider BIFF range"),
                Count::new(2).expect("count"),
                Count::new(1).expect("count"),
            ),
        )
        .expect_err("MS-OGRAPH rejects counts above 0x0F9F");
    assert!(matches!(
        error,
        Error::InvalidModel {
            field: "category count",
            ..
        }
    ));
    assert!(edit.is_empty());
    Ok(())
}

#[test]
fn physical_validation_is_failure_atomic() -> Result<()> {
    let source = fixture(DataKind::Text, 1, &[0xBB]);
    let original = source.as_bytes().to_vec();
    let mut chart = Chart::open(source, Context::graph())?;
    let scan = super::codec::scan(&chart)?;
    let invalid = Metadata::new(
        DataKind::Text,
        Count::new(0x0FA0).expect("Count permits the wider BIFF range"),
        Count::new(2).expect("count"),
        Count::new(1).expect("count"),
    );
    let change = super::Change::new(0, scan.entries[0].offset, scan.entries[0].metadata, invalid);
    assert!(matches!(
        super::codec::patch(&mut chart, scan.source, &[change]),
        Err(Error::InvalidModel {
            field: "category count",
            ..
        })
    ));
    assert_eq!(chart.encode()?.as_bytes(), original.as_slice());
    Ok(())
}
