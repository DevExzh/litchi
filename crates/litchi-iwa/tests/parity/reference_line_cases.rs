//! The reference-line graph engine (wave W1-E, record 0375a): a chart's
//! primary value-axis reference lines through the host editors and through
//! `litchi_iwa_style::reference_line::set_lines` over a
//! `litchi_iwa_style::graph` workspace of the same input — in-place patch,
//! removal (object, UUID binding, watermark release), allocation (object,
//! item UUID, component UUID, watermark) and the chart's
//! `reference_lines` extension. The chart is a column chart the host added
//! to a native-saved fixture.

use litchi_iwa::{keynote, numbers, pages};
use litchi_iwa_archive::iwa::Limits as CoreLimits;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::reference_line::{Line, Value};
use litchi_iwa_style::graph::Workspace;
use litchi_iwa_style::reference_line;

use crate::chart_style_cases::{Components, Target, input};
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier};
use crate::known_fields::Decoded;
use crate::series_graph_cases::publish_changes;

const CHART_MESSAGE_TYPE: u32 = 5_021;

/// The chart drawable with every reference-line item UUID zeroed (the one
/// documented divergence of an allocation: the host draws them at random,
/// the engine derives them), compared byte for byte otherwise.
fn chart_without_item_uuids(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    use litchi_iwa_protos::chart_style_codec::{self as codec, Limits};
    fn fail<E>(_: E) -> prost::DecodeError {
        // A truncated varint: the generated decoder's own error.
        <litchi_iwa_protos::tsp::Reference as prost::Message>::decode(&[0x08, 0xff][..])
            .expect_err("a truncated varint fails")
    }
    let normalised = match codec::reference_lines(data, Limits::DEFAULT).map_err(fail)? {
        Some(extension) => {
            let zero = [0x08, 0x00, 0x10, 0x00];
            let extension = litchi_iwa_common::wire::transform_length_delimited_fields_at_path(
                extension,
                &[1, 2],
                |item| -> Result<Vec<u8>, litchi_iwa_common::Error> {
                    litchi_iwa_common::wire::patch_length_delimited_field(
                        item,
                        2,
                        true,
                        Some(&zero),
                    )
                },
            )
            .map_err(fail)?;
            codec::rewrite_reference_lines(data, Some(&extension), Limits::DEFAULT)
                .map_err(fail)?
                .into_bytes()
                .into_owned()
        },
        None => data.to_vec(),
    };
    Ok(Decoded {
        debug: format!("{normalised:02x?}"),
        canonical: normalised,
        unknown_fields: std::collections::BTreeSet::new(),
    })
}

fn host_write(source: &[u8], target: Target, chart: u64, lines: &[Line]) -> TestResult<Vec<u8>> {
    Ok(match target {
        Target::Keynote { slide } => {
            let mut editor = keynote::Draft::from_bytes(source)?;
            editor.set_slide_chart_reference_lines(slide, chart, lines)?;
            editor.to_bytes()?
        },
        Target::Numbers { sheet } => {
            let mut editor = numbers::Draft::from_bytes(source)?;
            editor.set_sheet_chart_reference_lines(sheet, chart, lines)?;
            editor.to_bytes()?
        },
        Target::Pages => {
            let mut editor = pages::Draft::from_bytes(source)?;
            editor.set_body_chart_reference_lines(chart, lines)?;
            editor.to_bytes()?
        },
    })
}

fn host_read(bytes: &[u8], target: Target, chart: u64) -> TestResult<String> {
    Ok(match target {
        Target::Keynote { slide } => format!(
            "{:?}",
            keynote::Draft::from_bytes(bytes)?.slide_chart_reference_lines(slide, chart)?
        ),
        Target::Numbers { sheet } => format!(
            "{:?}",
            numbers::Draft::from_bytes(bytes)?.sheet_chart_reference_lines(sheet, chart)?
        ),
        Target::Pages => format!(
            "{:?}",
            pages::Draft::from_bytes(bytes)?.body_chart_reference_lines(chart)?
        ),
    })
}

fn engine_write(source: &[u8], chart: u64, lines: &[Line]) -> TestResult<Vec<u8>> {
    let components = Components::read(source)?;
    let mut workspace = Workspace::new(&components.archives, CoreLimits::default())?;
    reference_line::set_lines(&mut workspace, chart, lines)?;
    publish_changes(source, &components, workspace.finish())
}

/// Write `after`, first written with `before` through the host. A case
/// that allocates lines is tier B: the host draws the new items' UUIDs,
/// which live in the chart payload, at random; the engine derives them
/// deterministically, so the chart payload is compared byte for byte with
/// the item UUIDs zeroed.
fn line_case(
    name: &'static str,
    format: Format,
    tier: Tier,
    evidence: Evidence,
    before: &[Line],
    after: &[Line],
) -> TestResult {
    let mut input = input(format, Kind::Column2d)?;
    let (target, chart) = (input.target, input.chart);
    if !before.is_empty() {
        input.fixture.bytes = host_write(&input.fixture.bytes, target, chart, before)?;
    }
    let mut case = Case::new(name, format, tier, evidence, &input.fixture)
        .host(|source| host_write(source, target, chart, after))
        .leaf(|source| Ok(LeafRun::output(engine_write(source, chart, after)?)))
        .read("host reference-line read", move |bytes| {
            host_read(bytes, target, chart)
        });
    if tier == Tier::B {
        case = case.leaf_structural_metadata().known_fields(|known| {
            known.with(
                CHART_MESSAGE_TYPE,
                "TSCH.ChartDrawableArchive (item UUIDs zeroed)",
                chart_without_item_uuids,
            )
        });
    }
    case.run()?;
    Ok(())
}

fn named_average() -> TestResult<Line> {
    Ok(Line::average()
        .try_with_name("Mean")?
        .with_value_visibility(true))
}

fn threshold() -> TestResult<Line> {
    Ok(Line::custom(Value::new(42.5)?).try_with_name("Target")?)
}

#[test]
fn keynote_chart_reference_lines_allocate_parity() -> TestResult {
    line_case(
        "keynote_chart_reference_lines_allocate_parity",
        Format::Keynote,
        Tier::B,
        Evidence::Queued,
        &[],
        &[Line::average()],
    )
}

#[test]
fn numbers_chart_reference_lines_allocate_parity() -> TestResult {
    line_case(
        "numbers_chart_reference_lines_allocate_parity",
        Format::Numbers,
        Tier::B,
        Evidence::Queued,
        &[],
        &[threshold()?],
    )
}

#[test]
fn pages_chart_reference_lines_allocate_parity() -> TestResult {
    line_case(
        "pages_chart_reference_lines_allocate_parity",
        Format::Pages,
        Tier::B,
        Evidence::Queued,
        &[Line::minimum()],
        &[Line::minimum(), Line::maximum()],
    )
}

#[test]
fn keynote_chart_reference_lines_patch_parity() -> TestResult {
    line_case(
        "keynote_chart_reference_lines_patch_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        &[Line::average()],
        &[named_average()?],
    )
}

#[test]
fn numbers_chart_reference_lines_patch_parity() -> TestResult {
    line_case(
        "numbers_chart_reference_lines_patch_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        &[Line::median()],
        &[threshold()?],
    )
}

#[test]
fn pages_chart_reference_lines_remove_parity() -> TestResult {
    line_case(
        "pages_chart_reference_lines_remove_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &[Line::minimum(), Line::maximum()],
        &[Line::minimum()],
    )
}

#[test]
fn keynote_chart_reference_lines_remove_all_parity() -> TestResult {
    line_case(
        "keynote_chart_reference_lines_remove_all_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        &[Line::median()],
        &[],
    )
}

#[test]
fn numbers_chart_reference_lines_remove_all_parity() -> TestResult {
    line_case(
        "numbers_chart_reference_lines_remove_all_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        &[threshold()?, Line::maximum()],
        &[],
    )
}
