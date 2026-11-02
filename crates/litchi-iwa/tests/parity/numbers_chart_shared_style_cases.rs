//! Regressions (record 0414): a Numbers chart-level style write on a chart
//! style another chart also references must refuse, as the host's
//! `ChartStyleSlot::ensure_exclusive` does, instead of restyling both charts;
//! a chart shadow write refuses a series style another chart references (the
//! host's `ChartSeriesStyleSlot::ensure_exclusive`).

use litchi_iwa::charts::{ChartCornerRadius, ChartRoundedCorners, ChartShadow};
use litchi_iwa::numbers;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::chart::ChartData;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_protos::chart_style_codec::{self as codec, Limits};
use litchi_numbers::Package;
use litchi_numbers::sheet::chart::style::{Error, Property, Refusal, Shadow, corner};

use crate::chart_style_cases::{Components, Target, input, publish};
use crate::harness::{Format, TestResult};
use crate::leaf_bytes;

const CHART_MESSAGE_TYPE: u32 = 5_021;
/// `TSCH.ChartArchive.chart_style`, in the chart drawable's extension.
const CHART_STYLE_FIELD: u32 = 9;
/// `TSCH.ChartArchive.series_theme_styles`.
const SERIES_THEME_STYLES_FIELD: u32 = 17;

/// Two host-added charts, the second pointed at the first one's chart
/// style.
fn shared_source() -> TestResult<(Vec<u8>, u64, u64)> {
    shared_reference(CHART_STYLE_FIELD, |references| references.chart_style())
}

/// Two host-added charts, the second chart's first series theme style
/// pointed at the first chart's first one.
fn shared_series_source() -> TestResult<(Vec<u8>, u64, u64)> {
    shared_reference(SERIES_THEME_STYLES_FIELD, |references| {
        references.series_theme_styles().first().copied()
    })
}

/// Two host-added charts; the first `field` reference of the second chart
/// is pointed at the first chart's object `pick` selects.
fn shared_reference(
    field: u32,
    pick: impl Fn(&codec::StyleReferences) -> Option<u64>,
) -> TestResult<(Vec<u8>, u64, u64)> {
    let input = input(Format::Numbers, Kind::Column2d)?;
    let Target::Numbers { sheet } = input.target else {
        return Err("not a Numbers target".into());
    };
    let mut editor = numbers::Draft::from_bytes(&input.fixture.bytes)?;
    let second = editor.add_sheet_chart(
        sheet,
        Kind::Column2d,
        ChartData::new(
            vec!["A".to_owned(), "B".to_owned()],
            vec!["X".to_owned()],
            vec![vec![Some(1.0)], vec![Some(2.0)]],
        )?,
        DrawablePoint { x: 300.0, y: 140.0 },
        DrawableSize {
            width: 200.0,
            height: 160.0,
        },
    )?;
    let source = editor.to_bytes()?;
    let components = Components::read(&source)?;
    let first_location = components.locate(input.chart, CHART_MESSAGE_TYPE)?;
    let style = pick(
        &codec::references(
            components.payload(input.chart, first_location)?,
            Limits::DEFAULT,
        )
        .map_err(|error| error.to_string())?,
    )
    .ok_or("the first chart has no such style")?;
    let second_id = second.drawable_object_id;
    let second_location = components.locate(second_id, CHART_MESSAGE_TYPE)?;
    let mut first = true;
    let rewritten = litchi_iwa_common::wire::transform_length_delimited_fields_at_path(
        components.payload(second_id, second_location)?,
        &[10_000, field],
        |reference| {
            if std::mem::replace(&mut first, false) {
                litchi_iwa_common::wire::patch_varint_field(reference, 1, true, Some(style))
            } else {
                Ok(reference.to_vec())
            }
        },
    )?;
    let shared = publish(
        &source,
        components,
        vec![(second_id, CHART_MESSAGE_TYPE, rewritten)],
    )?;
    Ok((shared, sheet, input.chart))
}

#[test]
fn numbers_sheet_chart_style_shared_style_refuses() -> TestResult {
    let (source, sheet, chart) = shared_source()?;
    let rounding = corner::Rounding::new(corner::Radius::new(30.0)?, false);
    let mut host = numbers::Draft::from_bytes(&source)?;
    assert!(
        host.set_sheet_chart_rounded_corners(
            sheet,
            chart,
            ChartRoundedCorners::new(ChartCornerRadius::new(30.0)?, false),
        )
        .is_err(),
        "the host refuses a shared chart style"
    );
    let package = Package::from_bytes(&source)?;
    let before = leaf_bytes!(package);
    let mut edit = package.edit_sheet_chart_style(0, 1)?;
    edit.set(Property::RoundedCorners(rounding))?;
    assert!(matches!(
        edit.commit(),
        Err(Error::Refused(Refusal::SharedStyle))
    ));
    assert_eq!(leaf_bytes!(package), before);
    Ok(())
}

#[test]
fn numbers_sheet_chart_shadow_shared_series_style_refuses() -> TestResult {
    let (source, sheet, chart) = shared_series_source()?;
    let mut host = numbers::Draft::from_bytes(&source)?;
    assert!(
        host.set_sheet_chart_shadow(sheet, chart, ChartShadow::None)
            .is_err(),
        "the host refuses a shared series style"
    );
    let package = Package::from_bytes(&source)?;
    let before = leaf_bytes!(package);
    let mut edit = package.edit_sheet_chart_style(0, 1)?;
    edit.set(Property::Shadow(Shadow::None))?;
    assert!(matches!(
        edit.commit(),
        Err(Error::Refused(Refusal::SharedStyle))
    ));
    assert_eq!(leaf_bytes!(package), before);
    Ok(())
}
