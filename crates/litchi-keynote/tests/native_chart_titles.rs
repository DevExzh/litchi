//! Native Keynote integration coverage for chart and axis titles.
//!
//! The fixture was authored, saved, closed, and reopened by Keynote 14.4. It
//! contains one ordinary 2D column chart with chart, category-axis, and
//! value-axis titles. Physical owners are resolved from the generated IWA
//! messages and the selected lazy codecs, so this test remains independent of
//! native component names while requiring strict wire locality.

use std::{io, path::PathBuf};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::{keynote_chart_axis_title_codec, keynote_chart_title_codec, tsch};
use litchi_keynote::{Axis, ChartSelector, Package, SlideSelector};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const CHART_MESSAGE_TYPE: u32 = 5_021;
const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
const CHART_AXIS_MESSAGE_TYPE: u32 = 5_027;
const GENERATED_EXTENSION_FIELD: u32 = 10_000;
const SOURCE_CHART_TITLE: &str = "Native chart title — 北区";
const REPLACED_CHART_TITLE: &str = "Focused chart title — 北区";
const RESAVED_CHART_TITLE: &str = "Native saved chart title — 北区";
const SOURCE_VALUE_TITLE: &str = "Native revenue — 元";
const REPLACED_VALUE_TITLE: &str = "Focused revenue — 元";
const SOURCE_CATEGORY_TITLE: &str = "Native months — 月";
const REPLACED_CATEGORY_TITLE: &str = "Focused months — 月";
const RESAVED_VALUE_TITLE: &str = REPLACED_VALUE_TITLE;
const RESAVED_CATEGORY_TITLE: &str = REPLACED_CATEGORY_TITLE;
const SOURCE_CAPTION: &str = "Native chart caption marker";
const EXPECTED_ROWS: [&str; 2] = ["Region 1", "Region 2"];
const EXPECTED_COLUMNS: [&str; 4] = ["April", "May", "June", "July"];
const EXPECTED_VALUES: [[f64; 4]; 2] = [[17.0, 26.0, 53.0, 96.0], [55.0, 43.0, 70.0, 58.0]];
const ROOT_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/chart-titles-native.key")
}

fn resaved_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/chart-titles-native-resaved.key")
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
                    field.number() == GENERATED_EXTENSION_FIELD && field.wire_type() == 2
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

fn expected_chart_data(data: &ChartData) {
    assert_eq!(data, &ChartData::expected());
}

fn assert_chart_data_unchanged(source: &[u8], target: &[u8]) -> TestResult {
    let source_chart = chart_payload(source)?;
    let target_chart = chart_payload(target)?;
    assert_eq!(target_chart.member, source_chart.member);
    assert_eq!(target_chart.data, source_chart.data);
    expected_chart_data(&target_chart.data);
    Ok(())
}

fn preview_members(source: &[u8]) -> TestResult<Vec<String>> {
    Ok(Catalog::from_bytes(source)?
        .iter()
        .filter(|entry| ROOT_PREVIEWS.iter().any(|name| *name == entry.name()))
        .map(|entry| entry.name().to_owned())
        .collect())
}

fn title_owner_member(source: &[u8], expected: &str) -> TestResult<String> {
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
        let mut matched = false;
        for object in &archive.objects {
            for message in &object.messages {
                if message.type_ != CHART_NON_STYLE_MESSAGE_TYPE {
                    continue;
                }
                let outer = WireView::parse(&message.data)?;
                for field in outer.fields().filter(|field| {
                    field.number() == GENERATED_EXTENSION_FIELD && field.wire_type() == 2
                }) {
                    let title = keynote_chart_title_codec::decode_visible_chart_title(
                        field.payload(),
                        keynote_chart_title_codec::DecodeOptions::for_source(field.payload()),
                    )?;
                    matched |= title == Some(expected);
                }
            }
        }
        if matched && selected.replace(entry.name().to_owned()).is_some() {
            return Err(io::Error::other("native chart title owner is ambiguous").into());
        }
    }
    selected.ok_or_else(|| io::Error::other("native chart title owner is missing").into())
}

fn axis_owner_member(source: &[u8], axis: Axis, expected: &str) -> TestResult<String> {
    let catalog = Catalog::from_bytes(source)?;
    let kind = match axis {
        Axis::Category => keynote_chart_axis_title_codec::AxisTitleKind::Category,
        Axis::Value => keynote_chart_axis_title_codec::AxisTitleKind::Value,
    };
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
        let mut matched = false;
        for object in &archive.objects {
            for message in &object.messages {
                if message.type_ != CHART_AXIS_MESSAGE_TYPE {
                    continue;
                }
                let outer = WireView::parse(&message.data)?;
                for field in outer.fields().filter(|field| {
                    field.number() == GENERATED_EXTENSION_FIELD && field.wire_type() == 2
                }) {
                    let (snapshot, _) =
                        keynote_chart_axis_title_codec::decode_axis_titles_with_report(
                            field.payload(),
                            keynote_chart_axis_title_codec::DecodeOptions::for_source(
                                field.payload(),
                            ),
                        )?;
                    matched |= snapshot.visible_title(kind) == Some(expected);
                }
            }
        }
        if matched && selected.replace(entry.name().to_owned()).is_some() {
            return Err(io::Error::other("native chart axis owner is ambiguous").into());
        }
    }
    selected.ok_or_else(|| io::Error::other("native chart axis owner is missing").into())
}

fn assert_member_locality(source: &[u8], target: &[u8], changed_member: &str) -> TestResult {
    let source_catalog = Catalog::from_bytes(source)?;
    let target_catalog = Catalog::from_bytes(target)?;
    let previews = preview_members(source)?;
    let mut changed = false;
    let mut deleted = Vec::new();

    for source_entry in source_catalog.iter() {
        let Some(target_entry) = target_catalog
            .iter()
            .find(|entry| entry.name() == source_entry.name())
        else {
            deleted.push(source_entry.name().to_owned());
            continue;
        };
        if source_entry.name() == changed_member {
            assert_ne!(source_entry.data(), target_entry.data());
            changed = true;
        } else {
            assert_eq!(
                source_entry.data(),
                target_entry.data(),
                "unrelated member {} changed",
                source_entry.name()
            );
            assert_eq!(
                source_entry.raw_record().local_record(),
                target_entry.raw_record().local_record(),
                "unrelated member {} lost its exact local ZIP record",
                source_entry.name()
            );
        }
    }

    for target_entry in target_catalog.iter() {
        assert!(
            source_catalog
                .iter()
                .any(|entry| entry.name() == target_entry.name()),
            "title edit inserted an unexpected package member {}",
            target_entry.name()
        );
    }

    assert!(changed, "selected title owner was not rewritten");
    assert_eq!(deleted, previews, "changed title must remove root previews");
    Ok(())
}

fn assert_title(package: &Package, expected: Option<&str>) -> TestResult {
    assert_eq!(
        package.slide_chart_title(SlideSelector::index(0), ChartSelector::index(0))?,
        expected.map(str::to_owned)
    );
    Ok(())
}

fn assert_axis_title(package: &Package, axis: Axis, expected: Option<&str>) -> TestResult {
    assert_eq!(
        package.slide_chart_axis_title(SlideSelector::index(0), ChartSelector::index(0), axis,)?,
        expected.map(str::to_owned)
    );
    Ok(())
}

fn assert_native_caption_marker(package: &Package) -> TestResult {
    assert_eq!(
        package.slide_chart_caption(SlideSelector::index(0), ChartSelector::index(0))?,
        Some(SOURCE_CAPTION.to_owned())
    );
    Ok(())
}

#[test]
fn native_chart_titles_read_and_noop_exactly() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    expected_chart_data(&chart_payload(&source)?.data);
    assert_title(&package, Some(SOURCE_CHART_TITLE))?;
    assert_axis_title(&package, Axis::Value, Some(SOURCE_VALUE_TITLE))?;
    assert_axis_title(&package, Axis::Category, Some(SOURCE_CATEGORY_TITLE))?;
    assert_native_caption_marker(&package)?;
    assert_eq!(exact_bytes(&package)?, source);

    let chart_noop = package
        .edit_slide_chart_title(SlideSelector::index(0), ChartSelector::index(0))?
        .set(SOURCE_CHART_TITLE)?
        .commit()?;
    assert!(chart_noop.patch().is_noop());
    assert!(!chart_noop.diagnostics().changed());
    assert_eq!(chart_noop.diagnostics().touched_components(), 0);
    assert!(!chart_noop.diagnostics().full_reparse_performed());
    assert_eq!(exact_bytes(chart_noop.package())?, source);

    for (axis, title) in [
        (Axis::Value, SOURCE_VALUE_TITLE),
        (Axis::Category, SOURCE_CATEGORY_TITLE),
    ] {
        let noop = package
            .edit_slide_chart_axis_title(SlideSelector::index(0), ChartSelector::index(0), axis)?
            .set(title)?
            .commit()?;
        assert!(noop.patch().is_noop());
        assert!(!noop.diagnostics().changed());
        assert_eq!(noop.diagnostics().touched_components(), 0);
        assert_eq!(noop.diagnostics().deleted_previews(), 0);
        assert!(!noop.diagnostics().full_reparse_performed());
        assert_eq!(exact_bytes(noop.package())?, source);
    }
    Ok(())
}

#[test]
fn native_chart_title_set_clear_reopens_preserves_data_and_inverse() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let owner = title_owner_member(&source, SOURCE_CHART_TITLE)?;
    let previews = preview_members(&source)?;
    assert!(!previews.is_empty(), "native fixture has no root previews");

    let changed = package
        .edit_slide_chart_title(SlideSelector::index(0), ChartSelector::index(0))?
        .set(REPLACED_CHART_TITLE)?
        .commit()?;
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert_eq!(changed.diagnostics().deleted_previews(), previews.len());
    assert!(changed.diagnostics().full_reparse_performed());
    let changed_bytes = exact_bytes(changed.package())?;
    assert_title(changed.package(), Some(REPLACED_CHART_TITLE))?;
    assert_axis_title(changed.package(), Axis::Value, Some(SOURCE_VALUE_TITLE))?;
    assert_axis_title(
        changed.package(),
        Axis::Category,
        Some(SOURCE_CATEGORY_TITLE),
    )?;
    assert_native_caption_marker(changed.package())?;
    assert_chart_data_unchanged(&source, &changed_bytes)?;
    assert_member_locality(&source, &changed_bytes, &owner)?;
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert_eq!(
        Catalog::from_bytes(&changed_bytes)?.len(),
        Catalog::from_bytes(&source)?.len() - previews.len()
    );

    let reopened = Package::from_bytes(&changed_bytes)?;
    assert_title(&reopened, Some(REPLACED_CHART_TITLE))?;
    assert_chart_data_unchanged(&source, &changed_bytes)?;

    let restored = changed
        .package()
        .apply_slide_chart_title(&changed.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(restored.diagnostics().deleted_previews(), 0);
    assert_title(restored.package(), Some(SOURCE_CHART_TITLE))?;

    let forward_again = changed.patch().inverse().inverse();
    assert_eq!(forward_again, changed.patch().clone());
    let reapplied = package.apply_slide_chart_title(&forward_again)?;
    assert_eq!(exact_bytes(reapplied.package())?, changed_bytes);
    assert_eq!(reapplied.diagnostics().deleted_previews(), previews.len());

    let cleared = reopened
        .edit_slide_chart_title(SlideSelector::index(0), ChartSelector::index(0))?
        .clear()?
        .commit()?;
    let cleared_bytes = exact_bytes(cleared.package())?;
    assert_title(cleared.package(), None)?;
    assert_axis_title(cleared.package(), Axis::Value, Some(SOURCE_VALUE_TITLE))?;
    assert_axis_title(
        cleared.package(),
        Axis::Category,
        Some(SOURCE_CATEGORY_TITLE),
    )?;
    assert_chart_data_unchanged(&changed_bytes, &cleared_bytes)?;
    assert_member_locality(&changed_bytes, &cleared_bytes, &owner)?;
    let absent_clear = cleared
        .package()
        .edit_slide_chart_title(SlideSelector::index(0), ChartSelector::index(0))?
        .clear()?
        .commit()?;
    assert!(absent_clear.patch().is_noop());
    assert!(!absent_clear.diagnostics().changed());
    assert_eq!(exact_bytes(absent_clear.package())?, cleared_bytes);

    let restored_cleared = cleared
        .package()
        .apply_slide_chart_title(&cleared.patch().inverse())?;
    assert_eq!(exact_bytes(restored_cleared.package())?, changed_bytes);
    assert_title(restored_cleared.package(), Some(REPLACED_CHART_TITLE))?;
    Ok(())
}

#[test]
fn native_chart_axis_titles_set_clear_reopen_and_inverse_each_axis() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;

    for (axis, source_title, replacement) in [
        (Axis::Value, SOURCE_VALUE_TITLE, REPLACED_VALUE_TITLE),
        (
            Axis::Category,
            SOURCE_CATEGORY_TITLE,
            REPLACED_CATEGORY_TITLE,
        ),
    ] {
        let owner = axis_owner_member(&source, axis, source_title)?;
        let source_previews = preview_members(&source)?;
        assert!(
            !source_previews.is_empty(),
            "native fixture has no root previews"
        );
        let changed = package
            .edit_slide_chart_axis_title(SlideSelector::index(0), ChartSelector::index(0), axis)?
            .set(replacement)?
            .commit()?;
        assert!(changed.diagnostics().changed());
        assert_eq!(changed.diagnostics().touched_components(), 1);
        assert_eq!(
            changed.diagnostics().deleted_previews(),
            source_previews.len()
        );
        assert!(changed.diagnostics().full_reparse_performed());
        let changed_bytes = exact_bytes(changed.package())?;
        assert_axis_title(changed.package(), axis, Some(replacement))?;
        assert_title(changed.package(), Some(SOURCE_CHART_TITLE))?;
        assert_axis_title(
            changed.package(),
            match axis {
                Axis::Value => Axis::Category,
                Axis::Category => Axis::Value,
            },
            Some(match axis {
                Axis::Value => SOURCE_CATEGORY_TITLE,
                Axis::Category => SOURCE_VALUE_TITLE,
            }),
        )?;
        assert_native_caption_marker(changed.package())?;
        assert_chart_data_unchanged(&source, &changed_bytes)?;
        assert_member_locality(&source, &changed_bytes, &owner)?;

        let reopened = Package::from_bytes(&changed_bytes)?;
        assert_axis_title(&reopened, axis, Some(replacement))?;
        assert_chart_data_unchanged(&source, &changed_bytes)?;

        let restored = changed
            .package()
            .apply_slide_chart_axis_title(&changed.patch().inverse())?;
        assert_eq!(exact_bytes(restored.package())?, source);
        assert_eq!(restored.diagnostics().deleted_previews(), 0);
        assert_axis_title(restored.package(), axis, Some(source_title))?;

        let forward_again = changed.patch().inverse().inverse();
        assert_eq!(forward_again, changed.patch().clone());
        let reapplied = package.apply_slide_chart_axis_title(&forward_again)?;
        assert_eq!(exact_bytes(reapplied.package())?, changed_bytes);
        assert_eq!(
            reapplied.diagnostics().deleted_previews(),
            source_previews.len()
        );

        let cleared = reopened
            .edit_slide_chart_axis_title(SlideSelector::index(0), ChartSelector::index(0), axis)?
            .clear()?
            .commit()?;
        let cleared_bytes = exact_bytes(cleared.package())?;
        assert_axis_title(cleared.package(), axis, None)?;
        assert_title(cleared.package(), Some(SOURCE_CHART_TITLE))?;
        assert_chart_data_unchanged(&changed_bytes, &cleared_bytes)?;
        assert_member_locality(&changed_bytes, &cleared_bytes, &owner)?;
        assert_eq!(cleared.diagnostics().deleted_previews(), 0);

        let absent_clear = cleared
            .package()
            .edit_slide_chart_axis_title(SlideSelector::index(0), ChartSelector::index(0), axis)?
            .clear()?
            .commit()?;
        assert!(absent_clear.patch().is_noop());
        assert!(!absent_clear.diagnostics().changed());
        assert_eq!(exact_bytes(absent_clear.package())?, cleared_bytes);

        let restored_cleared = cleared
            .package()
            .apply_slide_chart_axis_title(&cleared.patch().inverse())?;
        assert_eq!(exact_bytes(restored_cleared.package())?, changed_bytes);
        assert_axis_title(restored_cleared.package(), axis, Some(replacement))?;
    }
    Ok(())
}

#[test]
fn native_resaved_chart_titles_read_and_noop_exactly() -> TestResult {
    let source = std::fs::read(resaved_fixture_path())?;
    let package = Package::from_bytes(&source)?;
    expected_chart_data(&chart_payload(&source)?.data);
    assert_title(&package, Some(RESAVED_CHART_TITLE))?;
    assert_axis_title(&package, Axis::Value, Some(RESAVED_VALUE_TITLE))?;
    assert_axis_title(&package, Axis::Category, Some(RESAVED_CATEGORY_TITLE))?;
    assert_native_caption_marker(&package)?;
    assert_eq!(exact_bytes(&package)?, source);

    let chart_noop = package
        .edit_slide_chart_title(SlideSelector::index(0), ChartSelector::index(0))?
        .set(RESAVED_CHART_TITLE)?
        .commit()?;
    assert!(chart_noop.patch().is_noop());
    assert!(!chart_noop.diagnostics().changed());
    assert_eq!(chart_noop.diagnostics().touched_components(), 0);
    assert_eq!(chart_noop.diagnostics().deleted_previews(), 0);
    assert!(!chart_noop.diagnostics().full_reparse_performed());
    assert_eq!(exact_bytes(chart_noop.package())?, source);

    for (axis, title) in [
        (Axis::Value, RESAVED_VALUE_TITLE),
        (Axis::Category, RESAVED_CATEGORY_TITLE),
    ] {
        let noop = package
            .edit_slide_chart_axis_title(SlideSelector::index(0), ChartSelector::index(0), axis)?
            .set(title)?
            .commit()?;
        assert!(noop.patch().is_noop());
        assert!(!noop.diagnostics().changed());
        assert_eq!(noop.diagnostics().touched_components(), 0);
        assert_eq!(noop.diagnostics().deleted_previews(), 0);
        assert!(!noop.diagnostics().full_reparse_performed());
        assert_eq!(exact_bytes(noop.package())?, source);
    }
    Ok(())
}
