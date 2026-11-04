//! A pie chart authored in Keynote 14.4 with only its third wedge exploded
//! (record 0433): six wedges, one override, stored `count` 3. The count is one
//! past the highest index, not the number of series or of entries.

use litchi_iwa_archive::Limits;
use litchi_iwa_archive::iwa::{Archive, SnappyStream};
use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::varint::decode_varint_from_bytes;
use litchi_iwa_common::wire::parse_wire_view;
use litchi_keynote::Package;
use litchi_keynote::chart::series::{Property, PropertyKind, WedgeExplosion};

const NATIVE: &[u8] = include_bytes!("../../../test-data/iwork/keynote/pie-third-wedge-native.key");
const CHART_DRAWABLE: u32 = 5_021;
const CHART_EXTENSION: u32 = 10_000;
const SERIES_NON_STYLES: u32 = 19;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

/// Field `number`'s payloads in `message`.
fn payloads(message: &[u8], number: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(parse_wire_view(message, WireLimits::DEFAULT)?
        .fields()
        .filter(|field| field.number() == number)
        .map(|field| field.payload().to_vec())
        .collect())
}

fn varint_of(message: &[u8], number: u32) -> TestResult<u64> {
    match payloads(message, number)?.first() {
        Some(payload) => Ok(decode_varint_from_bytes(payload)?.0),
        None => Ok(0),
    }
}

/// `(count, indexes)` of the chart's `series_non_styles`.
fn sparse_array(bytes: &[u8]) -> TestResult<(u64, Vec<u64>)> {
    let catalog = Catalog::from_bytes(bytes, Limits::default())?;
    for entry in catalog.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = SnappyStream::decompress(entry.data(), Default::default())?;
        let archive = Archive::parse(stream.as_bytes(), Default::default())?;
        for message in archive.objects.iter().flat_map(|object| &object.messages) {
            if message.type_ == CHART_DRAWABLE {
                let chart = payloads(&message.data, CHART_EXTENSION)?;
                let chart = chart.first().ok_or("the drawable has no chart")?;
                let array = payloads(chart, SERIES_NON_STYLES)?;
                let array = array.first().ok_or("no series array")?;
                let indexes = payloads(array, 2)?
                    .iter()
                    .map(|entry| varint_of(entry, 1))
                    .collect::<TestResult<Vec<_>>>()?;
                return Ok((varint_of(array, 1)?, indexes));
            }
        }
    }
    Err("no chart".into())
}

fn explosions(package: &Package) -> TestResult<Vec<WedgeExplosion>> {
    let Property::WedgeExplosions(values) =
        package.slide_chart_series(0usize, 0usize, PropertyKind::WedgeExplosions)?
    else {
        return Err("another property".into());
    };
    Ok(values)
}

#[test]
fn native_count_is_one_past_the_highest_index() -> TestResult {
    assert_eq!(sparse_array(NATIVE)?, (3, vec![2]));
    Ok(())
}

#[test]
fn leaf_reads_the_native_third_wedge() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    let expected = (0..6)
        .map(|wedge| WedgeExplosion::from_percent(if wedge == 2 { 25.0 } else { 0.0 }))
        .collect::<Result<Vec<_>, _>>()?;
    assert_eq!(explosions(&package)?, expected);
    Ok(())
}

/// The leaf's writer follows the native rule on Keynote's own chart: the fifth
/// wedge moves the end of the array, and the third keeps its value.
#[test]
fn leaf_writer_follows_the_native_count_rule() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    let values = (0..6)
        .map(|wedge| {
            WedgeExplosion::from_percent(match wedge {
                2 => 25.0,
                4 => 10.0,
                _ => 0.0,
            })
        })
        .collect::<Result<Vec<_>, _>>()?;
    let mut edit = package.edit_slide_chart_series(0usize, 0usize)?;
    edit.set(Property::WedgeExplosions(values.clone()))?;
    let commit = edit.commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    let (count, indexes) = sparse_array(&bytes)?;
    let highest = indexes.iter().copied().max().ok_or("no entry")?;
    assert_eq!(count, highest + 1);
    assert_eq!(highest, 4);
    assert_eq!(explosions(&Package::from_bytes(&bytes)?)?, values);
    Ok(())
}
