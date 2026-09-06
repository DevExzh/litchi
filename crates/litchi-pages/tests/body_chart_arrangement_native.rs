//! Native Pages coverage for the public body-chart Arrange-panel API.
//!
//! The fixtures are authored, saved, closed, and reopened by Pages.  The
//! test keeps the chart's decoded data grid as an independent semantic oracle
//! while checking that arrangement changes preserve every other package
//! member and every media asset.

use std::{io, path::PathBuf};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::tsch;
use litchi_pages::{BodyChartArrangementError, BodyChartSelector, ChartArrangement, Package};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const CHART_MESSAGE_TYPE: u32 = 5_021;
const CHART_EXTENSION_FIELD: u32 = 10_000;
const BODY_MARKER: &str = "Pages Arrange native body marker";
const EXPECTED_ROWS: [&str; 2] = ["Region 1", "Region 2"];
const EXPECTED_COLUMNS: [&str; 4] = ["April", "May", "June", "July"];
const EXPECTED_VALUES: [[f64; 4]; 2] = [[17.0, 26.0, 53.0, 96.0], [55.0, 43.0, 70.0, 58.0]];

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/chart-arrangement-native.pages")
}

fn retirement_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/chart-arrangement-retirement-resaved.pages")
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

#[derive(Debug, Clone, PartialEq)]
struct ChartData {
    /// The decoded TSCH archive is the semantic chart oracle.  Pages may keep
    /// a chart's visible grid inline or refer to a table-backed range; the
    /// complete archive handles both native layouts.
    archive: tsch::ChartArchive,
    /// Preserve the extension payload too, including fields outside the
    /// generated semantic projection.
    encoded: Vec<u8>,
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
            for message in &object.messages {
                if message.type_ != CHART_MESSAGE_TYPE {
                    continue;
                }
                let outer = WireView::parse(&message.data)?;
                for field in outer.fields().filter(|field| {
                    field.number() == CHART_EXTENSION_FIELD && field.wire_type() == 2
                }) {
                    let chart = tsch::ChartArchive::decode(field.payload())?;
                    if selected.is_some() {
                        return Err(io::Error::other(
                            "native Pages fixture contains multiple chart data payloads",
                        )
                        .into());
                    }
                    selected = Some(ChartPayload {
                        member: entry.name().to_owned(),
                        data: ChartData {
                            archive: chart,
                            encoded: field.payload().to_vec(),
                        },
                    });
                }
            }
        }
    }

    selected.ok_or_else(|| io::Error::other("native Pages chart data payload is missing").into())
}

fn assert_chart_data_unchanged(source: &[u8], target: &[u8]) -> TestResult {
    let source_chart = chart_payload(source)?;
    let target_chart = chart_payload(target)?;
    assert_eq!(target_chart.member, source_chart.member);
    assert_native_chart_data(&source_chart.data)?;
    assert_eq!(target_chart.data, source_chart.data);
    Ok(())
}

fn assert_native_chart_data(data: &ChartData) -> TestResult {
    let chart = &data.archive;
    assert_eq!(
        chart.chart_type,
        Some(tsch::ChartType::ColumnChartType2D as i32)
    );
    let grid = chart
        .grid
        .as_ref()
        .ok_or_else(|| io::Error::other("native Pages chart grid is missing"))?;
    assert_eq!(grid.row_name, EXPECTED_ROWS);
    assert_eq!(grid.column_name, EXPECTED_COLUMNS);
    let values = grid
        .grid_row
        .iter()
        .map(|row| {
            row.value
                .iter()
                .map(|value| value.numeric_value)
                .collect::<Vec<_>>()
        })
        .collect::<Vec<_>>();
    let expected = EXPECTED_VALUES
        .into_iter()
        .map(|row| row.into_iter().map(Some).collect::<Vec<_>>())
        .collect::<Vec<_>>();
    assert_eq!(values, expected);
    Ok(())
}

fn native_body_text(package: &Package) -> TestResult<String> {
    let text = package.text()?;
    assert!(text.contains(BODY_MARKER));
    assert_eq!(text.matches('\u{fffc}').count(), 1);
    Ok(text)
}

fn assert_member_locality(source: &[u8], target: &[u8], changed_member: &str) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    let mut changed = Vec::new();
    for entry in before.iter() {
        let candidate = after
            .iter()
            .find(|other| other.name() == entry.name())
            .ok_or_else(|| io::Error::other("Pages arrangement removed a package member"))?;
        if entry.name() == changed_member {
            assert_ne!(entry.data(), candidate.data());
            changed.push(entry.name().to_owned());
        } else {
            assert_eq!(
                entry.data(),
                candidate.data(),
                "unselected Pages package member {} changed",
                entry.name()
            );
            assert_eq!(
                entry.raw_record().local_record(),
                candidate.raw_record().local_record(),
                "unselected Pages package member {} lost its exact ZIP record",
                entry.name()
            );
        }
    }
    for entry in after.iter() {
        assert!(
            before.iter().any(|other| other.name() == entry.name()),
            "Pages arrangement inserted an unexpected package member {}",
            entry.name()
        );
    }
    changed.sort_unstable();
    assert_eq!(changed, [changed_member.to_owned()]);
    assert_eq!(before.len(), after.len());
    Ok(())
}

fn assert_data_assets_untouched(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    for entry in before
        .iter()
        .filter(|entry| entry.name().starts_with("Data/"))
    {
        let candidate = after
            .iter()
            .find(|other| other.name() == entry.name())
            .ok_or_else(|| io::Error::other("Pages arrangement removed a data asset"))?;
        assert_eq!(entry.data(), candidate.data(), "Pages data asset changed");
        assert_eq!(
            entry.raw_record().local_record(),
            candidate.raw_record().local_record(),
            "Pages data asset ZIP record changed"
        );
    }
    Ok(())
}

fn assert_arrangement(package: &Package, expected: ChartArrangement) -> TestResult {
    assert_eq!(
        package.body_chart_arrangement(BodyChartSelector::index(0))?,
        expected
    );
    Ok(())
}

#[test]
fn native_pages_chart_arrangement_reads_baseline_and_has_exact_noop() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let source_text = native_body_text(&package)?;
    assert_arrangement(&package, ChartArrangement::default())?;
    assert_chart_data_unchanged(&source, &source)?;
    assert_data_assets_untouched(&source, &source)?;
    assert_eq!(exact_bytes(&package)?, source);

    let no_op = package
        .edit_body_chart_arrangement(BodyChartSelector::index(0))?
        .set(ChartArrangement::default())
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());
    assert_eq!(exact_bytes(no_op.package())?, source);
    assert_eq!(no_op.package().text()?, source_text);
    Ok(())
}

#[test]
fn native_pages_chart_arrangement_changes_both_flags_reopens_and_inverts_exactly() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let source_text = native_body_text(&package)?;
    let location = chart_payload(&source)?;
    let baseline = ChartArrangement::default();
    let target = ChartArrangement::new(true, true);

    let changed = package
        .edit_body_chart_arrangement(BodyChartSelector::index(0))?
        .set(target)
        .commit()
        .map_err(|error| io::Error::other(format!("native Pages arrangement commit: {error:?}")))?;
    let candidate = exact_bytes(changed.package())?;
    assert_eq!(changed.patch().before(), baseline);
    assert_eq!(changed.patch().after(), target);
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_arrangement(changed.package(), target)?;
    assert_eq!(changed.package().text()?, source_text);
    assert_chart_data_unchanged(&source, &candidate)?;
    assert_data_assets_untouched(&source, &candidate)?;
    assert_member_locality(&source, &candidate, &location.member)?;

    let reopened = Package::from_bytes(&candidate)?;
    assert_arrangement(&reopened, target)?;
    assert_eq!(reopened.text()?, source_text);
    assert_chart_data_unchanged(&source, &candidate)?;

    let restored = changed
        .package()
        .apply_body_chart_arrangement(&changed.patch().inverse())
        .map_err(|error| {
            io::Error::other(format!("native Pages arrangement inverse apply: {error:?}"))
        })?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_arrangement(restored.package(), baseline)?;
    assert_eq!(restored.package().text()?, source_text);

    let forward_again = changed.patch().inverse().inverse();
    assert_eq!(forward_again, changed.patch().clone());
    let reapplied = package
        .apply_body_chart_arrangement(&forward_again)
        .map_err(|error| {
            io::Error::other(format!(
                "native Pages arrangement forward reapply: {error:?}"
            ))
        })?;
    assert_eq!(exact_bytes(reapplied.package())?, candidate);
    assert_arrangement(reapplied.package(), target)?;
    Ok(())
}

#[test]
fn native_pages_chart_arrangement_rejects_stale_patch_without_mutating_target() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let source_text = native_body_text(&package)?;
    let target = ChartArrangement::new(true, true);
    let changed = package
        .edit_body_chart_arrangement(BodyChartSelector::index(0))?
        .set(target)
        .commit()?;
    let candidate = exact_bytes(changed.package())?;
    let stale = changed
        .package()
        .apply_body_chart_arrangement(changed.patch())
        .expect_err("a patch must not apply to its own target as its source");
    assert!(matches!(stale, BodyChartArrangementError::PatchConflict));
    assert_eq!(exact_bytes(changed.package())?, candidate);
    assert_arrangement(changed.package(), target)?;
    assert_eq!(changed.package().text()?, source_text);
    Ok(())
}

#[test]
fn native_pages_chart_arrangement_retirement_oracle_reads_true_state_and_resets_exactly()
-> TestResult {
    let source = std::fs::read(retirement_fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let source_text = native_body_text(&package)?;
    let location = chart_payload(&source)?;
    let native_state = ChartArrangement::new(true, true);

    assert_arrangement(&package, native_state)?;
    assert_chart_data_unchanged(&source, &source)?;
    assert_data_assets_untouched(&source, &source)?;
    assert_eq!(exact_bytes(&package)?, source);

    let no_op = package
        .edit_body_chart_arrangement(BodyChartSelector::index(0))?
        .set(native_state)
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());
    assert_eq!(exact_bytes(no_op.package())?, source);
    assert_eq!(no_op.package().text()?, source_text);

    let reset = package
        .edit_body_chart_arrangement(BodyChartSelector::index(0))?
        .set(ChartArrangement::default())
        .commit()?;
    let reset_bytes = exact_bytes(reset.package())?;
    assert_arrangement(reset.package(), ChartArrangement::default())?;
    assert_eq!(reset.package().text()?, source_text);
    assert_chart_data_unchanged(&source, &reset_bytes)?;
    assert_data_assets_untouched(&source, &reset_bytes)?;
    assert_member_locality(&source, &reset_bytes, &location.member)?;

    let reopened = Package::from_bytes(&reset_bytes)?;
    assert_arrangement(&reopened, ChartArrangement::default())?;
    assert_eq!(reopened.text()?, source_text);
    let restored = reset
        .package()
        .apply_body_chart_arrangement(&reset.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_arrangement(restored.package(), native_state)?;
    assert_eq!(restored.package().text()?, source_text);
    Ok(())
}
