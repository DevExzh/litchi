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

/// A host-added chart with a host-set legend font, and a second chart that
/// also consumes the legend's paragraph-style variation (its
/// `paragraph_styles` list and metadata name it). The host's
/// `style_is_exclusive_to_chart` then sees another consumer, so a font
/// change replaces the shared variation (`replace_shared_style`).
fn shared_legend_variation() -> TestResult<(crate::harness::Fixture, u64, u64)> {
    use litchi_iwa::charts::{ChartFont, ChartLegendFont};
    use litchi_iwa::text::TextFont;
    let input = input(Format::Numbers, Kind::Column2d)?;
    let Target::Numbers { sheet } = input.target else {
        return Err("not a Numbers target".into());
    };
    let mut editor = numbers::Draft::from_bytes(&input.fixture.bytes)?;
    editor.set_sheet_chart_legend_font(
        sheet,
        input.chart,
        &ChartLegendFont::Font(ChartFont::new(TextFont::named("Georgia")?).with_bold(true)),
    )?;
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
    let first = components.locate(input.chart, CHART_MESSAGE_TYPE)?;
    let variation = *codec::references(components.payload(input.chart, first)?, Limits::DEFAULT)
        .map_err(|error| error.to_string())?
        .paragraph_styles()
        .last()
        .ok_or("the first chart has no paragraph styles")?;
    let second_id = second.drawable_object_id;
    let second_location = components.locate(second_id, CHART_MESSAGE_TYPE)?;
    let mut reference = Vec::new();
    litchi_iwa_common::wire::append_varint_field(&mut reference, 1, variation)?;
    let rewritten = litchi_iwa_common::wire::transform_length_delimited_field(
        components.payload(second_id, second_location)?,
        10_000,
        |chart| {
            let mut chart = chart.to_vec();
            litchi_iwa_common::wire::append_length_delimited_field(
                &mut chart,
                20,
                &reference,
                litchi_iwa_common::WireLimits::DEFAULT,
            )?;
            Ok::<_, litchi_iwa_common::Error>(chart)
        },
    )?;
    let shared = crate::chart_style_cases::publish_referencing(
        &source,
        components,
        vec![(second_id, CHART_MESSAGE_TYPE, rewritten, vec![variation])],
    )?;
    Ok((
        crate::harness::Fixture {
            path: "chart-data-native.numbers + shared legend variation",
            bytes: shared,
        },
        sheet,
        input.chart,
    ))
}

/// The host replaces a legend paragraph-style variation another chart also
/// consumes; the leaf takes the same route (record 0414 §13).
#[test]
fn numbers_sheet_chart_legend_font_shared_variation_parity() -> TestResult {
    use crate::harness::{Case, Evidence, LeafRun, Tier};
    use litchi_iwa::charts::{ChartFont, ChartLegendFont, ChartLegendFontSize};
    use litchi_iwa::text::TextFont;
    use litchi_numbers::sheet::chart::legend_style::{
        Font, LegendFont, LegendFontSize, Property as LegendProperty, PropertyKind, TextPointSize,
    };
    let (fixture, sheet, chart) = shared_legend_variation()?;
    // The fixture really shares the variation: the host's font change
    // replaces it in the first chart's list.
    let styles = |bytes: &[u8]| -> TestResult<Vec<u64>> {
        let components = Components::read(bytes)?;
        let location = components.locate(chart, CHART_MESSAGE_TYPE)?;
        Ok(
            codec::references(components.payload(chart, location)?, Limits::DEFAULT)
                .map_err(|error| error.to_string())?
                .paragraph_styles()
                .to_vec(),
        )
    };
    let mut host = numbers::Draft::from_bytes(&fixture.bytes)?;
    host.set_sheet_chart_legend_font(
        sheet,
        chart,
        &ChartLegendFont::Font(ChartFont::new(TextFont::named("Courier New")?)),
    )?;
    let before = styles(&fixture.bytes)?;
    let after = styles(&host.to_bytes()?)?;
    assert_eq!(before.len(), after.len());
    assert_ne!(
        before.last(),
        after.last(),
        "the shared variation is replaced"
    );
    let cases = [
        (
            LegendProperty::Font(LegendFont::Font {
                font: Font::named("Courier New")?,
                bold: false,
                italic: true,
            }),
            false,
        ),
        (
            LegendProperty::FontSize(LegendFontSize::Size(TextPointSize::from_points(20.0)?)),
            false,
        ),
        (LegendProperty::Font(LegendFont::Inherited), true),
    ];
    for (value, _) in cases {
        let host_value = value.clone();
        let staged = value.clone();
        let kind = value.kind();
        Case::new(
            "numbers_sheet_chart_legend_font_shared_variation_parity",
            Format::Numbers,
            Tier::B,
            Evidence::Queued,
            &fixture,
        )
        .host(|source| {
            let mut e = numbers::Draft::from_bytes(source)?;
            match &host_value {
                LegendProperty::Font(LegendFont::Font { font, bold, italic }) => e
                    .set_sheet_chart_legend_font(
                        sheet,
                        chart,
                        &ChartLegendFont::Font(
                            ChartFont::new(font.clone())
                                .with_bold(*bold)
                                .with_italic(*italic),
                        ),
                    )?,
                LegendProperty::Font(_) => {
                    e.set_sheet_chart_legend_font(sheet, chart, &ChartLegendFont::Inherited)?;
                },
                LegendProperty::FontSize(LegendFontSize::Size(size)) => e
                    .set_sheet_chart_legend_font_size(
                        sheet,
                        chart,
                        ChartLegendFontSize::Size((*size).into()),
                    )?,
                _ => return Err("unmapped".into()),
            }
            Ok(e.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_sheet_chart_legend_style(0, 1)?;
            edit.set(staged.clone())?;
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_sheet_chart_legend_style(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host legend typography", move |bytes| {
            let e = numbers::Draft::from_bytes(bytes)?;
            Ok(if kind == PropertyKind::Font {
                format!("{:?}", e.sheet_chart_legend_font(sheet, chart)?)
            } else {
                format!("{:?}", e.sheet_chart_legend_font_size(sheet, chart)?)
            })
        })
        .leaf_structural_metadata()
        .run()
        .map_err(|error| format!("{value:?}: {error}"))?;
    }
    Ok(())
}
