//! Host-vs-leaf chart title and axis-title parity for W1-E slice 0375b.
//!
//! The host keeps the differential oracle: each case mutates one selector
//! route, reopens both outputs through the leaf, compares the semantic read,
//! and asks the leaf to apply the inverse patch. Caption authoring is
//! intentionally absent from this slice.

use litchi_iwa::numbers;
use litchi_iwa::pages;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::chart::ChartData;
use litchi_iwa_common::chart::axis::Axis;
use litchi_iwa_common::chart::kind::Kind;
use litchi_numbers::{ChartSelector, Package as NumbersPackage, SheetSelector};
use litchi_pages::{BodyChartSelector, Package as PagesPackage};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const NUMBERS: &str = "iwork/numbers/chart-data-native.numbers";
/// The native Numbers chart is followed by the host-created chart used by
/// these mutation vectors. This matches the established chart-style parity
/// setup; the arrangement fixture has a stale host UUID map.
const NUMBERS_CHART: usize = 1;
const PAGES_CHART: usize = 0;

fn numbers_input() -> TestResult<(Fixture, u64, u64)> {
    let native = fixture(NUMBERS)?;
    let mut editor = numbers::Draft::from_bytes(&native.bytes)?;
    let sheet = editor
        .sheets()?
        .first()
        .ok_or("native Numbers sheet missing")?
        .id();
    let data = ChartData::new(
        vec!["Region 1".to_owned(), "Region 2".to_owned()],
        vec!["April".to_owned(), "May".to_owned(), "June".to_owned()],
        vec![
            vec![Some(17.0), Some(26.0), Some(53.0)],
            vec![Some(55.0), Some(43.0), Some(70.0)],
        ],
    )?;
    let chart = editor.add_sheet_chart(
        sheet,
        Kind::Column2d,
        data,
        DrawablePoint { x: 120.0, y: 140.0 },
        DrawableSize {
            width: 360.0,
            height: 240.0,
        },
    )?;
    Ok((
        Fixture {
            path: "chart-data-native.numbers + add_sheet_chart",
            bytes: editor.to_bytes()?,
        },
        sheet,
        chart.drawable_object_id,
    ))
}

#[derive(Clone, Copy, Debug)]
enum Property {
    Title,
    Axis(Axis),
}

fn pages_input() -> TestResult<(Fixture, u64)> {
    let mut editor = pages::Draft::create_with_text("W1-E chart text parity")?;
    let chart = editor
        .add_body_chart(
            "W1-E chart text parity".encode_utf16().count(),
            Kind::Column2d,
            ChartData::new(
                vec!["Region 1".to_owned(), "Region 2".to_owned()],
                vec!["April".to_owned(), "May".to_owned(), "June".to_owned()],
                vec![
                    vec![Some(17.0), Some(26.0), Some(53.0)],
                    vec![Some(55.0), Some(43.0), Some(70.0)],
                ],
            )?,
            DrawablePoint { x: 120.0, y: 140.0 },
            DrawableSize {
                width: 360.0,
                height: 240.0,
            },
        )?
        .drawable_object_id;
    Ok((
        Fixture {
            path: "source-built: Pages chart title and axis parity",
            bytes: editor.to_bytes()?,
        },
        chart,
    ))
}

fn host_numbers(source: &[u8], sheet: u64, chart: u64, property: Property) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    match property {
        Property::Title => editor.set_sheet_chart_title(sheet, chart, "W1-E title")?,
        Property::Axis(axis) => {
            editor.set_sheet_chart_axis_title(sheet, chart, axis, "W1-E axis title")?;
        },
    }
    Ok(editor.to_bytes()?)
}

fn leaf_numbers(source: &[u8], property: Property) -> TestResult<LeafRun> {
    let package = NumbersPackage::from_bytes(source)?;
    let (output, restored) = match property {
        Property::Title => {
            let commit = package
                .edit_sheet_chart_title(
                    SheetSelector::index(0),
                    ChartSelector::index(NUMBERS_CHART),
                )?
                .set("W1-E title")?
                .commit()?;
            let restored = commit
                .package()
                .apply_sheet_chart_title(&commit.patch().inverse())?;
            (
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
        Property::Axis(axis) => {
            let commit = package
                .edit_sheet_chart_axis_title(
                    SheetSelector::index(0),
                    ChartSelector::index(NUMBERS_CHART),
                    axis,
                )?
                .set("W1-E axis title")?
                .commit()?;
            let restored = commit
                .package()
                .apply_sheet_chart_axis_title(&commit.patch().inverse())?;
            (
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
    };
    Ok(LeafRun::with_inverse(output, restored))
}

fn read_numbers(source: &[u8], property: Property) -> TestResult<Option<String>> {
    let package = NumbersPackage::from_bytes(source)?;
    Ok(match property {
        Property::Title => package
            .sheet_chart_title(SheetSelector::index(0), ChartSelector::index(NUMBERS_CHART))?,
        Property::Axis(axis) => package.sheet_chart_axis_title(
            SheetSelector::index(0),
            ChartSelector::index(NUMBERS_CHART),
            axis,
        )?,
    })
}

fn host_read_numbers(
    source: &[u8],
    sheet: u64,
    chart: u64,
    property: Property,
) -> TestResult<Option<String>> {
    let editor = numbers::Draft::from_bytes(source)?;
    Ok(match property {
        Property::Title => editor.sheet_chart_title(sheet, chart)?,
        Property::Axis(axis) => editor.sheet_chart_axis_title(sheet, chart, axis)?,
    })
}

fn host_pages(source: &[u8], chart: u64, property: Property) -> TestResult<Vec<u8>> {
    let mut editor = pages::Draft::from_bytes(source)?;
    match property {
        Property::Title => editor.set_body_chart_title(chart, "W1-E title")?,
        Property::Axis(axis) => editor.set_body_chart_axis_title(chart, axis, "W1-E axis title")?,
    }
    Ok(editor.to_bytes()?)
}

fn leaf_pages(source: &[u8], property: Property) -> TestResult<LeafRun> {
    let package = PagesPackage::from_bytes(source)?;
    let (output, restored) = match property {
        Property::Title => {
            let commit = package
                .edit_body_chart_title(BodyChartSelector::index(PAGES_CHART))?
                .set("W1-E title")?
                .commit()?;
            let restored = commit
                .package()
                .apply_body_chart_title(&commit.patch().inverse())?;
            (
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
        Property::Axis(axis) => {
            let commit = package
                .edit_body_chart_axis_title(BodyChartSelector::index(PAGES_CHART), axis)?
                .set("W1-E axis title")?
                .commit()?;
            let restored = commit
                .package()
                .apply_body_chart_axis_title(&commit.patch().inverse())?;
            (
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            )
        },
    };
    Ok(LeafRun::with_inverse(output, restored))
}

fn read_pages(source: &[u8], property: Property) -> TestResult<Option<String>> {
    let package = PagesPackage::from_bytes(source)?;
    Ok(match property {
        Property::Title => package.body_chart_title(BodyChartSelector::index(PAGES_CHART))?,
        Property::Axis(axis) => {
            package.body_chart_axis_title(BodyChartSelector::index(PAGES_CHART), axis)?
        },
    })
}

fn host_read_pages(source: &[u8], chart: u64, property: Property) -> TestResult<Option<String>> {
    let editor = pages::Draft::from_bytes(source)?;
    Ok(match property {
        Property::Title => editor.body_chart_title(chart)?,
        Property::Axis(axis) => editor.body_chart_axis_title(chart, axis)?,
    })
}

fn assert_read_equal(label: &str, host: &Option<String>, leaf: &Option<String>) -> TestResult {
    if host == leaf {
        Ok(())
    } else {
        Err(format!("{label}: host {host:?} != leaf {leaf:?}").into())
    }
}

fn mismatched_read(value: &Option<String>) -> Option<String> {
    match value {
        Some(value) => Some(format!("{value} [mismatched read]")),
        None => Some("W1-E mismatched read".to_owned()),
    }
}

fn host_numbers_remove(
    source: &[u8],
    sheet: u64,
    chart: u64,
    property: Property,
) -> TestResult<Vec<u8>> {
    // The caller has already authored the common populated input through the
    // host setter.  This route is intentionally clear-only so the
    // differential case proves the same operation on the same source.
    let mut editor = numbers::Draft::from_bytes(source)?;
    match property {
        Property::Title => {
            let _ = editor.remove_sheet_chart_title(sheet, chart)?;
        },
        Property::Axis(axis) => {
            let _ = editor.remove_sheet_chart_axis_title(sheet, chart, axis)?;
        },
    }
    Ok(editor.to_bytes()?)
}

fn leaf_numbers_remove(source: &[u8], property: Property) -> TestResult<LeafRun> {
    // Match host_numbers_remove: clear the caller-provided populated source,
    // then apply that exact commit's inverse to prove restoration.
    let package = NumbersPackage::from_bytes(source)?;
    match property {
        Property::Title => {
            let commit = package
                .edit_sheet_chart_title(
                    SheetSelector::index(0),
                    ChartSelector::index(NUMBERS_CHART),
                )?
                .clear()?
                .commit()?;
            let restored = commit
                .package()
                .apply_sheet_chart_title(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        },
        Property::Axis(axis) => {
            let commit = package
                .edit_sheet_chart_axis_title(
                    SheetSelector::index(0),
                    ChartSelector::index(NUMBERS_CHART),
                    axis,
                )?
                .clear()?
                .commit()?;
            let restored = commit
                .package()
                .apply_sheet_chart_axis_title(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        },
    }
}

fn host_pages_remove(source: &[u8], chart: u64, property: Property) -> TestResult<Vec<u8>> {
    // The caller has already authored the common populated input through the
    // host setter; this route only clears the selected property.
    let mut editor = pages::Draft::from_bytes(source)?;
    match property {
        Property::Title => {
            let _ = editor.remove_body_chart_title(chart)?;
        },
        Property::Axis(axis) => {
            let _ = editor.remove_body_chart_axis_title(chart, axis)?;
        },
    }
    Ok(editor.to_bytes()?)
}

fn leaf_pages_remove(source: &[u8], property: Property) -> TestResult<LeafRun> {
    // Match host_pages_remove: clear the caller-provided populated source and
    // verify the exact inverse restores it.
    let package = PagesPackage::from_bytes(source)?;
    match property {
        Property::Title => {
            let commit = package
                .edit_body_chart_title(BodyChartSelector::index(PAGES_CHART))?
                .clear()?
                .commit()?;
            let restored = commit
                .package()
                .apply_body_chart_title(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        },
        Property::Axis(axis) => {
            let commit = package
                .edit_body_chart_axis_title(BodyChartSelector::index(PAGES_CHART), axis)?
                .clear()?
                .commit()?;
            let restored = commit
                .package()
                .apply_body_chart_axis_title(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        },
    }
}

fn run_numbers_reads(input: &Fixture, sheet: u64, chart: u64, property: Property) -> TestResult {
    let source = input.bytes.clone();
    let host_read = host_read_numbers(&source, sheet, chart, property)?;
    let leaf_read = read_numbers(&source, property)?;
    assert_read_equal("Numbers direct chart-text read", &host_read, &leaf_read)?;
    assert_eq!(
        source, input.bytes,
        "chart-text read changed the source bytes"
    );
    Ok(())
}

fn run_pages_reads(input: &Fixture, chart: u64, property: Property) -> TestResult {
    let source = input.bytes.clone();
    let host_read = host_read_pages(&source, chart, property)?;
    let leaf_read = read_pages(&source, property)?;
    assert_read_equal("Pages direct chart-text read", &host_read, &leaf_read)?;
    assert_eq!(
        source, input.bytes,
        "chart-text read changed the source bytes"
    );
    Ok(())
}

#[test]
fn chart_text_read_negative_control_rejects_mismatched_values() -> TestResult {
    let (numbers, sheet, chart) = numbers_input()?;
    let (pages, pages_chart) = pages_input()?;
    let values = vec![
        (
            "Numbers title",
            host_read_numbers(&numbers.bytes, sheet, chart, Property::Title)?,
        ),
        (
            "Numbers category axis",
            host_read_numbers(&numbers.bytes, sheet, chart, Property::Axis(Axis::Category))?,
        ),
        (
            "Numbers value axis",
            host_read_numbers(&numbers.bytes, sheet, chart, Property::Axis(Axis::Value))?,
        ),
        (
            "Pages title",
            host_read_pages(&pages.bytes, pages_chart, Property::Title)?,
        ),
        (
            "Pages category axis",
            host_read_pages(&pages.bytes, pages_chart, Property::Axis(Axis::Category))?,
        ),
        (
            "Pages value axis",
            host_read_pages(&pages.bytes, pages_chart, Property::Axis(Axis::Value))?,
        ),
        ("present value", Some("W1-E title".to_owned())),
    ];
    for (label, value) in values {
        let mismatch = mismatched_read(&value);
        if assert_read_equal(label, &value, &mismatch).is_ok() {
            return Err(format!("chart-text read comparison accepted a mismatched {label}").into());
        }
    }
    Ok(())
}

#[test]
fn numbers_chart_title_and_axis_title_set_parity() -> TestResult {
    let (input, sheet, chart) = numbers_input()?;
    for property in [
        Property::Title,
        Property::Axis(Axis::Category),
        Property::Axis(Axis::Value),
    ] {
        Case::new(
            "numbers_chart_text_title_axis_set_parity",
            Format::Numbers,
            Tier::A,
            Evidence::Queued,
            &input,
        )
        .host(move |source| host_numbers(source, sheet, chart, property))
        .leaf(|source| leaf_numbers(source, property))
        .read("numbers chart text", move |source| {
            read_numbers(source, property)
        })
        .run()?;
    }
    Ok(())
}

#[test]
fn numbers_chart_title_and_axis_title_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    let (input, sheet, chart) = numbers_input()?;
    for property in [
        Property::Title,
        Property::Axis(Axis::Category),
        Property::Axis(Axis::Value),
    ] {
        run_numbers_reads(&input, sheet, chart, property)?;
    }
    Ok(())
}

#[test]
fn numbers_chart_title_and_axis_title_remove_parity() -> TestResult {
    let (mut input, sheet, chart) = numbers_input()?;
    for property in [
        Property::Title,
        Property::Axis(Axis::Category),
        Property::Axis(Axis::Value),
    ] {
        input.bytes = host_numbers(&input.bytes, sheet, chart, property)?;
        Case::new(
            "numbers_chart_text_title_axis_remove_parity",
            Format::Numbers,
            Tier::B,
            Evidence::Queued,
            &input,
        )
        .host(move |source| host_numbers_remove(source, sheet, chart, property))
        .leaf(|source| leaf_numbers_remove(source, property))
        .read("numbers chart text", move |source| {
            read_numbers(source, property)
        })
        .run()?;
    }
    Ok(())
}

#[test]
fn pages_chart_title_and_axis_title_set_parity() -> TestResult {
    let (input, chart) = pages_input()?;
    for property in [
        Property::Title,
        Property::Axis(Axis::Category),
        Property::Axis(Axis::Value),
    ] {
        Case::new(
            "pages_chart_text_title_axis_set_parity",
            Format::Pages,
            Tier::A,
            Evidence::Queued,
            &input,
        )
        .host(move |source| host_pages(source, chart, property))
        .leaf(|source| leaf_pages(source, property))
        .read("pages chart text", move |source| {
            read_pages(source, property)
        })
        .run()?;
    }
    Ok(())
}

#[test]
fn pages_chart_title_and_axis_title_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    let (input, chart) = pages_input()?;
    for property in [
        Property::Title,
        Property::Axis(Axis::Category),
        Property::Axis(Axis::Value),
    ] {
        run_pages_reads(&input, chart, property)?;
    }
    Ok(())
}

#[test]
fn pages_chart_title_and_axis_title_remove_parity() -> TestResult {
    let (mut input, chart) = pages_input()?;
    for property in [
        Property::Title,
        Property::Axis(Axis::Category),
        Property::Axis(Axis::Value),
    ] {
        input.bytes = host_pages(&input.bytes, chart, property)?;
        Case::new(
            "pages_chart_text_title_axis_remove_parity",
            Format::Pages,
            Tier::B,
            Evidence::Queued,
            &input,
        )
        .host(move |source| host_pages_remove(source, chart, property))
        .leaf(|source| leaf_pages_remove(source, property))
        .read("pages chart text", move |source| {
            read_pages(source, property)
        })
        .run()?;
    }
    Ok(())
}
