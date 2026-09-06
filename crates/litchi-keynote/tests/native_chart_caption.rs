//! Native Keynote integration coverage for chart-caption transactions.
//!
//! The fixture was authored, saved, closed, and reopened by Keynote 14.4.  It
//! contains one ordinary 2D column chart.  This test resolves the physical
//! chart and metadata members from their typed IWA messages instead of baking
//! native component names into the test, then checks that caption edits leave
//! the chart data and unrelated ZIP records untouched.

use std::{io, path::PathBuf};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::{pages_movie_caption_codec, tsch};
use litchi_keynote::{ChartSelector, Package, SlideSelector};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const CHART_MESSAGE_TYPE: u32 = 5_021;
const CHART_EXTENSION_FIELD: u32 = 10_000;
const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const SOURCE_CAPTION: &str = "Native chart caption marker";
const FOCUSED_CAPTION: &str = "Focused chart caption — 北区";
const RESAVED_CAPTION: &str = "Native saved caption — 北区";
const EXPECTED_ROWS: [&str; 2] = ["Region 1", "Region 2"];
const EXPECTED_COLUMNS: [&str; 4] = ["April", "May", "June", "July"];
const EXPECTED_VALUES: [[f64; 4]; 2] = [[17.0, 26.0, 53.0, 96.0], [55.0, 43.0, 70.0, 58.0]];

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/chart-caption-native.key")
}

fn resaved_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/chart-caption-native-resaved.key")
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

#[derive(Debug, Clone, PartialEq)]
struct ChartData {
    chart_type: Option<i32>,
    rows: Vec<String>,
    columns: Vec<String>,
    values: Vec<Vec<Option<f64>>>,
}

impl ChartData {
    fn from_archive(chart: &tsch::ChartArchive) -> TestResult<Self> {
        let grid = chart
            .grid
            .as_ref()
            .ok_or_else(|| io::Error::other("native chart has no inline data grid"))?;
        Ok(Self {
            chart_type: chart.chart_type,
            rows: grid.row_name.clone(),
            columns: grid.column_name.clone(),
            values: grid
                .grid_row
                .iter()
                .map(|row| row.value.iter().map(|value| value.numeric_value).collect())
                .collect(),
        })
    }

    fn expected() -> Self {
        Self {
            chart_type: Some(tsch::ChartType::ColumnChartType2D as i32),
            rows: EXPECTED_ROWS
                .iter()
                .map(|value| (*value).to_owned())
                .collect(),
            columns: EXPECTED_COLUMNS
                .iter()
                .map(|value| (*value).to_owned())
                .collect(),
            values: EXPECTED_VALUES
                .into_iter()
                .map(|row| row.into_iter().map(Some).collect())
                .collect(),
        }
    }
}

#[derive(Debug, Clone)]
struct ChartPayload {
    member: String,
    data: ChartData,
}

fn chart_payload(source: &[u8]) -> TestResult<ChartPayload> {
    let catalog = Catalog::from_bytes(source)?;
    let mut selected = None;

    for entry in catalog.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = match SnappyStream::decompress(entry.data()) {
            Ok(stream) => stream,
            Err(_) => continue,
        };
        let archive = match Archive::parse(stream.as_bytes()) {
            Ok(archive) => archive,
            Err(_) => continue,
        };
        for object in &archive.objects {
            if object.archive_info.identifier.is_none() {
                return Err(io::Error::other("native chart object has no identifier").into());
            }
            for message in &object.messages {
                if message.type_ != CHART_MESSAGE_TYPE {
                    continue;
                }
                let outer = WireView::parse(&message.data)?;
                for field in outer.fields().filter(|field| {
                    field.number() == CHART_EXTENSION_FIELD && field.wire_type() == 2
                }) {
                    let chart = tsch::ChartArchive::decode(field.payload())?;
                    if chart.grid.is_none() {
                        continue;
                    }
                    let data = ChartData::from_archive(&chart)?;
                    if selected.is_some() {
                        return Err(io::Error::other(
                            "native fixture contains multiple matching chart payloads",
                        )
                        .into());
                    }
                    selected = Some(ChartPayload {
                        member: entry.name().to_owned(),
                        data,
                    });
                }
            }
        }
    }

    selected.ok_or_else(|| io::Error::other("native column chart payload is missing").into())
}

fn metadata_member(source: &[u8]) -> TestResult<String> {
    let catalog = Catalog::from_bytes(source)?;
    let mut selected = None;
    for entry in catalog.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = match SnappyStream::decompress(entry.data()) {
            Ok(stream) => stream,
            Err(_) => continue,
        };
        let archive = match Archive::parse(stream.as_bytes()) {
            Ok(archive) => archive,
            Err(_) => continue,
        };
        let count = archive
            .objects
            .iter()
            .flat_map(|object| object.messages.iter())
            .filter(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
            .count();
        if count == 0 {
            continue;
        }
        if count != 1 || selected.replace(entry.name().to_owned()).is_some() {
            return Err(io::Error::other("native fixture has ambiguous package metadata").into());
        }
    }
    selected.ok_or_else(|| io::Error::other("native package metadata member is missing").into())
}

fn slide_node_member(source: &[u8]) -> TestResult<String> {
    let catalog = Catalog::from_bytes(source)?;
    let mut selected = catalog
        .iter()
        .filter(|entry| entry.name().rsplit('/').next() == Some("Document.iwa"));
    let entry = selected
        .next()
        .ok_or_else(|| io::Error::other("native root Document.iwa member is missing"))?;
    if selected.next().is_some() {
        return Err(io::Error::other("native fixture has multiple Document.iwa members").into());
    }
    Ok(entry.name().to_owned())
}

fn caption_style_member(source: &[u8]) -> TestResult<String> {
    let catalog = Catalog::from_bytes(source)?;
    let mut style_identifier = None;
    for entry in catalog.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = match SnappyStream::decompress(entry.data()) {
            Ok(stream) => stream,
            Err(_) => continue,
        };
        let archive = match Archive::parse(stream.as_bytes()) {
            Ok(archive) => archive,
            Err(_) => continue,
        };
        for object in &archive.objects {
            for message in &object.messages {
                if message.type_ != CAPTION_INFO_MESSAGE_TYPE {
                    continue;
                }
                let snapshot = pages_movie_caption_codec::decode_caption_info(
                    &message.data,
                    pages_movie_caption_codec::DecodeOptions::for_source(&message.data),
                )?;
                let identifier = snapshot
                    .style_identifier()
                    .ok_or_else(|| io::Error::other("native caption has no style reference"))?;
                if style_identifier
                    .replace(identifier)
                    .is_some_and(|previous| previous != identifier)
                {
                    return Err(
                        io::Error::other("native fixture has multiple caption styles").into(),
                    );
                }
            }
        }
    }
    let style_identifier = style_identifier
        .ok_or_else(|| io::Error::other("native caption style reference is missing"))?;
    let mut member = None;
    for entry in catalog.iter() {
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = match SnappyStream::decompress(entry.data()) {
            Ok(stream) => stream,
            Err(_) => continue,
        };
        let archive = match Archive::parse(stream.as_bytes()) {
            Ok(archive) => archive,
            Err(_) => continue,
        };
        for object in &archive.objects {
            if object.archive_info.identifier != Some(style_identifier) {
                continue;
            }
            let style_messages = object
                .messages
                .iter()
                .filter(|message| message.type_ == SHAPE_STYLE_MESSAGE_TYPE)
                .count();
            if style_messages != 1 || member.replace(entry.name().to_owned()).is_some() {
                return Err(io::Error::other("native caption style owner is ambiguous").into());
            }
        }
    }
    member.ok_or_else(|| io::Error::other("native caption style member is missing").into())
}

fn preview_members(source: &[u8]) -> TestResult<Vec<String>> {
    Ok(Catalog::from_bytes(source)?
        .iter()
        .filter(|entry| {
            !entry.name().ends_with(".iwa")
                && entry
                    .name()
                    .rsplit('/')
                    .next()
                    .is_some_and(|name| name.starts_with("preview"))
        })
        .map(|entry| entry.name().to_owned())
        .collect())
}

fn assert_expected_chart_data(data: &ChartData) {
    assert_eq!(data, &ChartData::expected());
}

fn assert_chart_data_unchanged(source: &[u8], target: &[u8]) -> TestResult {
    let source_chart = chart_payload(source)?;
    let target_chart = chart_payload(target)?;
    assert_eq!(target_chart.data, source_chart.data);
    assert_expected_chart_data(&target_chart.data);
    Ok(())
}

fn assert_caption_locality(
    source: &[u8],
    target: &[u8],
    chart_member: &str,
    rewrites_slide_node: bool,
) -> TestResult {
    let source_catalog = Catalog::from_bytes(source)?;
    let target_catalog = Catalog::from_bytes(target)?;
    let metadata = metadata_member(source)?;
    let previews = preview_members(source)?;
    let mut changed = Vec::new();
    let mut deleted = Vec::new();

    for source_entry in source_catalog.iter() {
        let Some(target_entry) = target_catalog
            .iter()
            .find(|entry| entry.name() == source_entry.name())
        else {
            deleted.push(source_entry.name().to_owned());
            continue;
        };
        if source_entry.data() != target_entry.data() {
            changed.push(source_entry.name().to_owned());
        } else {
            assert_eq!(
                source_entry.raw_record().local_record(),
                target_entry.raw_record().local_record(),
                "unchanged member {} lost its exact local ZIP record",
                source_entry.name()
            );
        }
    }

    for target_entry in target_catalog.iter() {
        assert!(
            source_catalog
                .iter()
                .any(|entry| entry.name() == target_entry.name()),
            "caption edit inserted an unexpected package member {}",
            target_entry.name()
        );
    }

    changed.sort_unstable();
    deleted.sort_unstable();
    let mut expected_changed = vec![chart_member.to_owned(), metadata];
    if rewrites_slide_node {
        expected_changed.push(slide_node_member(source)?);
    }
    expected_changed.sort_unstable();
    let mut expected_deleted = previews;
    expected_deleted.sort_unstable();
    assert_eq!(changed, expected_changed);
    assert_eq!(deleted, expected_deleted);
    assert_eq!(
        target_catalog.len(),
        source_catalog.len().saturating_sub(expected_deleted.len())
    );
    Ok(())
}

fn assert_caption_style_untouched(source: &[u8], target: &[u8]) -> TestResult {
    let member = caption_style_member(source)?;
    let source_catalog = Catalog::from_bytes(source)?;
    let target_catalog = Catalog::from_bytes(target)?;
    let source_entry = source_catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other("native caption style member is missing"))?;
    let target_entry = target_catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other("caption edit removed the style member"))?;
    assert_eq!(target_entry.data(), source_entry.data());
    assert_eq!(
        target_entry.raw_record().local_record(),
        source_entry.raw_record().local_record(),
        "external caption style lost its exact local ZIP record"
    );
    Ok(())
}

fn assert_caption(package: &Package, expected: Option<&str>) -> TestResult {
    assert_eq!(
        package.slide_chart_caption(SlideSelector::index(0), ChartSelector::index(0))?,
        expected.map(str::to_owned)
    );
    Ok(())
}

#[test]
fn native_chart_caption_reads_data_and_has_an_exact_noop() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let location = chart_payload(&source)?;
    assert_expected_chart_data(&location.data);
    assert_caption(&package, Some(SOURCE_CAPTION))?;
    assert_eq!(exact_bytes(&package)?, source);

    let no_op = package
        .edit_slide_chart_caption(SlideSelector::index(0), ChartSelector::index(0))?
        .set(SOURCE_CAPTION)?
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());
    assert_eq!(exact_bytes(no_op.package())?, source);
    Ok(())
}

#[test]
fn native_chart_caption_set_reopens_preserves_data_and_inverse() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let location = chart_payload(&source)?;
    let changed = package
        .edit_slide_chart_caption(SlideSelector::index(0), ChartSelector::index(0))?
        .set(FOCUSED_CAPTION)?
        .commit()?;
    let candidate = exact_bytes(changed.package())?;

    assert_caption(changed.package(), Some(FOCUSED_CAPTION))?;
    assert_chart_data_unchanged(&source, &candidate)?;
    assert_caption_locality(&source, &candidate, &location.member, true)?;
    assert_caption_style_untouched(&source, &candidate)?;

    let reopened = Package::from_bytes(&candidate)?;
    assert_caption(&reopened, Some(FOCUSED_CAPTION))?;
    assert_chart_data_unchanged(&source, &candidate)?;

    let restored = changed
        .package()
        .apply_slide_chart_caption(&changed.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_caption(restored.package(), Some(SOURCE_CAPTION))?;
    Ok(())
}

#[test]
fn native_chart_caption_clear_is_reversible_and_absent_clear_is_exact_noop() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let location = chart_payload(&source)?;
    let cleared = package
        .edit_slide_chart_caption(SlideSelector::index(0), ChartSelector::index(0))?
        .clear()?
        .commit()?;
    let cleared_bytes = exact_bytes(cleared.package())?;
    assert_caption(cleared.package(), None)?;
    assert_chart_data_unchanged(&source, &cleared_bytes)?;
    assert_caption_locality(&source, &cleared_bytes, &location.member, false)?;
    assert_caption_style_untouched(&source, &cleared_bytes)?;

    let absent_clear = cleared
        .package()
        .edit_slide_chart_caption(SlideSelector::index(0), ChartSelector::index(0))?
        .clear()?
        .commit()?;
    assert!(absent_clear.patch().is_noop());
    assert!(!absent_clear.diagnostics().changed());
    assert_eq!(exact_bytes(absent_clear.package())?, cleared_bytes);

    let restored = cleared
        .package()
        .apply_slide_chart_caption(&cleared.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_caption(restored.package(), Some(SOURCE_CAPTION))?;
    Ok(())
}

#[test]
fn native_resaved_chart_caption_reads_data_and_has_an_exact_noop() -> TestResult {
    let source = std::fs::read(resaved_fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let location = chart_payload(&source)?;
    assert_expected_chart_data(&location.data);
    assert_caption(&package, Some(RESAVED_CAPTION))?;
    assert_eq!(exact_bytes(&package)?, source);
    assert_caption_style_untouched(&source, &source)?;

    let no_op = package
        .edit_slide_chart_caption(SlideSelector::index(0), ChartSelector::index(0))?
        .set(RESAVED_CAPTION)?
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert!(!no_op.diagnostics().changed());
    assert_eq!(exact_bytes(no_op.package())?, source);
    Ok(())
}
