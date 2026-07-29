//! Source-built Keynote chart CRUD regression tests.

use std::fs;
use std::path::PathBuf;

use super::*;
use crate::charts::{
    ChartAxis, ChartAxisBound, ChartAxisMajorStepCount, ChartAxisMinorStepCount,
    ChartAxisTickMarkLocation, ChartCornerRadius, ChartDonutInnerRadius, ChartErrorBarDirection,
    ChartErrorBarFixedValue, ChartErrorBarPercentage, ChartGapPercentage, ChartGapSpacing,
    ChartLegendFill, ChartPieLabelDistance, ChartPieLabelVisibility, ChartPieStartAngle,
    ChartPieWedgeExplosion, ChartPieWedgeIndex, ChartRoundedCorners, ChartSeriesErrorBarAutoFit,
    ChartSeriesErrorBars, ChartSeriesIndex, ChartSeriesStroke, ChartSeriesStrokePattern,
    ChartSeriesTrendline, ChartSeriesTrendlineMovingAveragePeriod,
    ChartSeriesTrendlinePolynomialOrder, ChartSeriesValueLabelAffixes,
    ChartSeriesValueLabelAutoFit, ChartSeriesValueLabelDecimalPlaces,
    ChartSeriesValueLabelLocation, ChartSeriesValueLabelNegativeStyle,
    ChartSeriesValueLabelNumberFormat, ChartSeriesValueLabelVisibility, ChartShadow,
    ChartValueAxisBounds, ChartValueAxisScale, ChartValueAxisSteps,
};
use crate::keynote::KeynoteDocumentBuilder;
use crate::shapes::{
    RgbColorSpace, RgbaColor, ShapeDropShadow, ShapeFill, ShapeImageFillTechnique,
    ShapeShadowAngle, ShapeShadowAppearance, ShapeShadowBlurRadius, ShapeShadowOffset,
    ShapeShadowOpacity, ShapeStroke, StrokePattern, StrokeWidth,
};

const POSITION: DrawablePoint = DrawablePoint { x: 240.0, y: 260.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 720.0,
    height: 420.0,
};

fn sample_data() -> ChartData {
    ChartData::new(
        vec!["Region 1".to_owned(), "Region 2".to_owned()],
        vec![
            "April".to_owned(),
            "May".to_owned(),
            "June".to_owned(),
            "July".to_owned(),
        ],
        vec![
            vec![Some(17.0), Some(26.0), Some(53.0), Some(96.0)],
            vec![Some(55.0), Some(43.0), Some(70.0), Some(58.0)],
        ],
    )
    .unwrap()
}

fn pie_data() -> ChartData {
    ChartData::new(
        vec!["North".to_owned(), "South".to_owned(), "West".to_owned()],
        vec!["Revenue".to_owned()],
        vec![vec![Some(12.0)], vec![Some(18.0)], vec![Some(24.0)]],
    )
    .unwrap()
}

fn gap_spacing(between_items: f32, between_sets: f32) -> ChartGapSpacing {
    ChartGapSpacing::new(
        ChartGapPercentage::new(between_items).unwrap(),
        ChartGapPercentage::new(between_sets).unwrap(),
    )
}

fn chart_stroke(pattern: StrokePattern, width: f32) -> ShapeStroke {
    ShapeStroke::new(
        RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
        StrokeWidth::new(width).unwrap(),
        pattern,
    )
}

fn chart_series_stroke(pattern: ChartSeriesStrokePattern, width: f32) -> ChartSeriesStroke {
    ChartSeriesStroke::new(
        RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
        StrokeWidth::new(width).unwrap(),
        pattern,
    )
}

fn chart_background_fill() -> ShapeFill {
    ShapeFill::Solid(RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap())
}

fn chart_shadow() -> ChartShadow {
    ChartShadow::Grouped(ShapeDropShadow::new(
        ShapeShadowAppearance::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
            ShapeShadowBlurRadius::from_points(15).unwrap(),
            ShapeShadowOffset::from_points(8.0).unwrap(),
            ShapeShadowOpacity::new(0.6).unwrap(),
        ),
        ShapeShadowAngle::from_degrees(60.0).unwrap(),
    ))
}

fn fixture(relative: &str) -> Vec<u8> {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    fs::read(root.join(relative)).unwrap()
}

#[test]
fn scratch_presentation_supports_standalone_chart_crud() {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Source-built chart")
        .build()
        .unwrap();
    let baseline = editor.to_bytes().unwrap();

    let created = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    assert_eq!(created.kind, ChartKind::Column2d);
    assert_eq!(created.direction, ChartSeriesDirection::Rows);
    assert_eq!(created.data, sample_data());

    let replacement = ChartData::new(
        vec!["Revenue".to_owned()],
        vec!["2026".to_owned(), "2027".to_owned(), "2028".to_owned()],
        vec![vec![Some(30.0), Some(45.0), None]],
    )
    .unwrap();
    editor
        .set_slide_chart_kind(0, created.drawable_object_id, ChartKind::Bar2d)
        .unwrap();
    editor
        .set_slide_chart_data(0, created.drawable_object_id, replacement.clone())
        .unwrap();
    editor
        .set_slide_chart_direction(0, created.drawable_object_id, ChartSeriesDirection::Columns)
        .unwrap();
    let changed_geometry = chart_geometry(
        "Keynote",
        DrawablePoint { x: 360.0, y: 180.0 },
        DrawableSize {
            width: 840.0,
            height: 480.0,
        },
    )
    .unwrap();
    editor
        .set_slide_chart_geometry(0, created.drawable_object_id, changed_geometry)
        .unwrap();

    let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let chart = &reopened.slide_charts(0).unwrap()[0];
    assert_eq!(chart.kind, ChartKind::Bar2d);
    assert_eq!(chart.direction, ChartSeriesDirection::Columns);
    assert_eq!(chart.data, replacement);
    assert_eq!(chart.geometry, changed_geometry);

    let removed = editor
        .remove_slide_chart(0, created.drawable_object_id)
        .unwrap();
    assert_eq!(removed.chart.drawable_object_id, created.drawable_object_id);
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn chart_creation_rejects_invalid_inputs_transactionally() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let baseline = editor.to_bytes().unwrap();
    assert!(
        editor
            .add_slide_chart(0, ChartKind::Undefined, sample_data(), POSITION, SIZE)
            .is_err()
    );
    assert!(
        editor
            .add_slide_chart(
                0,
                ChartKind::Column2d,
                sample_data(),
                POSITION,
                DrawableSize {
                    width: 0.0,
                    height: SIZE.height,
                },
            )
            .is_err()
    );
    assert!(
        editor
            .add_slide_chart(1, ChartKind::Column2d, sample_data(), POSITION, SIZE)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn multiple_chart_theme_registrations_are_removed_independently() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let baseline = editor.to_bytes().unwrap();
    let first = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let second = editor
        .add_slide_chart(
            0,
            ChartKind::Line2d,
            sample_data(),
            DrawablePoint {
                x: POSITION.x + SIZE.width,
                y: POSITION.y,
            },
            SIZE,
        )
        .unwrap();

    editor
        .remove_slide_chart(0, first.drawable_object_id)
        .unwrap();
    assert_eq!(
        editor
            .slide_charts(0)
            .unwrap()
            .iter()
            .map(|chart| chart.drawable_object_id)
            .collect::<Vec<_>>(),
        vec![second.drawable_object_id]
    );
    editor
        .remove_slide_chart(0, second.drawable_object_id)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn duplicate_slide_chart_clones_the_private_graph_and_inline_data() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let source_graph = chart_graph(&editor, 0, source.drawable_object_id).unwrap();
    let baseline = editor.to_bytes().unwrap();
    assert!(editor.duplicate_slide_chart(0, u64::MAX).is_err());
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    let duplicate_graph = chart_graph(&editor, 0, duplicate.drawable_object_id).unwrap();
    let expected_geometry =
        offset_drawable_geometry(source.geometry, DRAWABLE_DUPLICATE_OFFSET).unwrap();

    assert_ne!(duplicate.drawable_object_id, source.drawable_object_id);
    assert_eq!(duplicate.kind, source.kind);
    assert_eq!(duplicate.direction, source.direction);
    assert_eq!(duplicate.data, source.data);
    assert_eq!(duplicate.geometry, expected_geometry);
    assert_eq!(
        duplicate_graph.object_ids.len(),
        source_graph.object_ids.len()
    );
    assert!(
        source_graph
            .object_ids
            .iter()
            .all(|identifier| !duplicate_graph.object_ids.contains(identifier))
    );

    let replacement = ChartData::new(
        vec!["Revenue".to_owned()],
        vec!["2026".to_owned(), "2027".to_owned()],
        vec![vec![Some(30.0), Some(45.0)]],
    )
    .unwrap();
    editor
        .set_slide_chart_data(0, duplicate.drawable_object_id, replacement.clone())
        .unwrap();
    assert_eq!(
        chart_graph(&editor, 0, source.drawable_object_id)
            .unwrap()
            .info
            .data,
        source.data
    );
    assert_eq!(
        chart_graph(&editor, 0, duplicate.drawable_object_id)
            .unwrap()
            .info
            .data,
        replacement
    );

    editor
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(
        editor
            .slide_charts(0)
            .unwrap()
            .iter()
            .map(|chart| chart.drawable_object_id)
            .collect::<Vec<_>>(),
        vec![duplicate.drawable_object_id]
    );
    editor
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(editor.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_chart_caption_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();

    assert_eq!(
        editor
            .slide_chart_caption(0, source.drawable_object_id)
            .unwrap(),
        None
    );
    editor
        .set_slide_chart_caption(0, source.drawable_object_id, "Revenue by region")
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_caption(0, source.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_caption(0, duplicate.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );

    editor
        .set_slide_chart_caption(0, source.drawable_object_id, "Updated source caption")
        .unwrap();
    assert!(
        editor
            .remove_slide_chart_caption(0, source.drawable_object_id)
            .unwrap()
    );
    assert!(
        !editor
            .remove_slide_chart_caption(0, source.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        editor
            .slide_chart_caption(0, source.drawable_object_id)
            .unwrap(),
        None
    );

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_caption(0, duplicate.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(
        reopened
            .slide_charts(0)
            .unwrap()
            .iter()
            .all(|chart| chart.drawable_object_id != duplicate.drawable_object_id)
    );
}

#[test]
fn scratch_presentation_supports_native_chart_title_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();

    assert_eq!(
        editor
            .slide_chart_title(0, source.drawable_object_id)
            .unwrap(),
        None
    );
    editor
        .set_slide_chart_title(0, source.drawable_object_id, "Revenue by region")
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_title(0, source.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_title(0, duplicate.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );

    editor
        .set_slide_chart_title(0, source.drawable_object_id, "Updated source title")
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_title(0, source.drawable_object_id)
            .unwrap(),
        Some("Updated source title".to_owned())
    );
    assert_eq!(
        editor
            .slide_chart_title(0, duplicate.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );
    assert!(
        editor
            .remove_slide_chart_title(0, source.drawable_object_id)
            .unwrap()
    );
    assert!(
        !editor
            .remove_slide_chart_title(0, source.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        editor
            .slide_chart_title(0, source.drawable_object_id)
            .unwrap(),
        None
    );

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_title(0, duplicate.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(
        reopened
            .slide_charts(0)
            .unwrap()
            .iter()
            .all(|chart| chart.drawable_object_id != duplicate.drawable_object_id)
    );
}

#[test]
fn scratch_presentation_supports_native_chart_axis_title_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();

    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert_eq!(
            editor
                .slide_chart_axis_title(0, source.drawable_object_id, axis)
                .unwrap(),
            None
        );
    }
    editor
        .set_slide_chart_axis_title(0, source.drawable_object_id, ChartAxis::Category, "Month")
        .unwrap();
    editor
        .set_slide_chart_axis_title(0, source.drawable_object_id, ChartAxis::Value, "Revenue")
        .unwrap();

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    for (axis, title) in [
        (ChartAxis::Category, "Month"),
        (ChartAxis::Value, "Revenue"),
    ] {
        assert_eq!(
            editor
                .slide_chart_axis_title(0, source.drawable_object_id, axis)
                .unwrap()
                .as_deref(),
            Some(title)
        );
        assert_eq!(
            editor
                .slide_chart_axis_title(0, duplicate.drawable_object_id, axis)
                .unwrap()
                .as_deref(),
            Some(title)
        );
    }

    editor
        .set_slide_chart_axis_title(
            0,
            source.drawable_object_id,
            ChartAxis::Category,
            "Updated month",
        )
        .unwrap();
    editor
        .set_slide_chart_axis_title(
            0,
            source.drawable_object_id,
            ChartAxis::Value,
            "Updated revenue",
        )
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_axis_title(0, source.drawable_object_id, ChartAxis::Category)
            .unwrap()
            .as_deref(),
        Some("Updated month")
    );
    assert_eq!(
        editor
            .slide_chart_axis_title(0, source.drawable_object_id, ChartAxis::Value)
            .unwrap()
            .as_deref(),
        Some("Updated revenue")
    );
    assert_eq!(
        editor
            .slide_chart_axis_title(0, duplicate.drawable_object_id, ChartAxis::Category)
            .unwrap()
            .as_deref(),
        Some("Month")
    );
    assert_eq!(
        editor
            .slide_chart_axis_title(0, duplicate.drawable_object_id, ChartAxis::Value)
            .unwrap()
            .as_deref(),
        Some("Revenue")
    );

    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert!(
            editor
                .remove_slide_chart_axis_title(0, source.drawable_object_id, axis)
                .unwrap()
        );
        assert!(
            !editor
                .remove_slide_chart_axis_title(0, source.drawable_object_id, axis)
                .unwrap()
        );
    }

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_axis_title(0, duplicate.drawable_object_id, ChartAxis::Category)
            .unwrap()
            .as_deref(),
        Some("Month")
    );
    assert_eq!(
        reopened
            .slide_chart_axis_title(0, duplicate.drawable_object_id, ChartAxis::Value)
            .unwrap()
            .as_deref(),
        Some("Revenue")
    );
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(
        reopened
            .slide_charts(0)
            .unwrap()
            .iter()
            .all(|chart| chart.drawable_object_id != duplicate.drawable_object_id)
    );
}

#[test]
fn scratch_presentation_supports_native_chart_value_axis_bounds_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let automatic = ChartValueAxisBounds::automatic();
    let fixed = ChartValueAxisBounds::fixed(
        ChartAxisBound::new(-10.0).unwrap(),
        ChartAxisBound::new(40.0).unwrap(),
    )
    .unwrap();
    let minimum_only =
        ChartValueAxisBounds::new(Some(ChartAxisBound::new(-5.0).unwrap()), None).unwrap();

    assert_eq!(
        editor
            .slide_chart_value_axis_bounds(0, source.drawable_object_id)
            .unwrap(),
        automatic
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_value_axis_bounds(0, source.drawable_object_id, automatic)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_value_axis_bounds(0, source.drawable_object_id, fixed)
        .unwrap();
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_value_axis_bounds(0, duplicate.drawable_object_id)
            .unwrap(),
        fixed
    );

    editor
        .set_slide_chart_value_axis_bounds(0, source.drawable_object_id, minimum_only)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_value_axis_bounds(0, source.drawable_object_id)
            .unwrap(),
        minimum_only
    );

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_value_axis_bounds(0, source.drawable_object_id)
            .unwrap(),
        minimum_only
    );
    assert_eq!(
        reopened
            .slide_chart_value_axis_bounds(0, duplicate.drawable_object_id)
            .unwrap(),
        fixed
    );
    reopened
        .set_slide_chart_value_axis_bounds(0, source.drawable_object_id, automatic)
        .unwrap();
    assert_eq!(
        reopened
            .slide_chart_value_axis_bounds(0, source.drawable_object_id)
            .unwrap(),
        automatic
    );
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_presentation_supports_native_chart_border_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();

    assert!(
        !editor
            .slide_chart_border_visible(0, source.drawable_object_id)
            .unwrap()
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_border_visible(0, source.drawable_object_id, false)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_border_visible(0, source.drawable_object_id, true)
        .unwrap();
    assert!(
        editor
            .slide_chart_border_visible(0, source.drawable_object_id)
            .unwrap()
    );

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert!(
        editor
            .slide_chart_border_visible(0, duplicate.drawable_object_id)
            .unwrap()
    );
    editor
        .set_slide_chart_border_visible(0, source.drawable_object_id, false)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert!(
        !reopened
            .slide_chart_border_visible(0, source.drawable_object_id)
            .unwrap()
    );
    assert!(
        reopened
            .slide_chart_border_visible(0, duplicate.drawable_object_id)
            .unwrap()
    );
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_chart_rounded_corner_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let rounded = ChartRoundedCorners::new(ChartCornerRadius::new(20.0).unwrap(), true);
    let changed = ChartRoundedCorners::new(ChartCornerRadius::new(35.0).unwrap(), false);

    assert_eq!(
        editor
            .slide_chart_rounded_corners(0, source.drawable_object_id)
            .unwrap(),
        ChartRoundedCorners::NONE
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_rounded_corners(0, source.drawable_object_id, ChartRoundedCorners::NONE)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_rounded_corners(0, source.drawable_object_id, rounded)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_rounded_corners(0, source.drawable_object_id)
            .unwrap(),
        rounded
    );

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_rounded_corners(0, duplicate.drawable_object_id)
            .unwrap(),
        rounded
    );
    editor
        .set_slide_chart_rounded_corners(0, source.drawable_object_id, changed)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_rounded_corners(0, source.drawable_object_id)
            .unwrap(),
        changed
    );
    assert_eq!(
        reopened
            .slide_chart_rounded_corners(0, duplicate.drawable_object_id)
            .unwrap(),
        rounded
    );
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_chart_gap_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let customized = gap_spacing(25.0, 70.0);
    let changed = gap_spacing(30.0, 60.0);

    assert_eq!(
        editor
            .slide_chart_gap_spacing(0, source.drawable_object_id)
            .unwrap(),
        ChartGapSpacing::NATIVE_DEFAULT
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_gap_spacing(
            0,
            source.drawable_object_id,
            ChartGapSpacing::NATIVE_DEFAULT,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_gap_spacing(0, source.drawable_object_id, customized)
        .unwrap();
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_gap_spacing(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    editor
        .set_slide_chart_gap_spacing(0, source.drawable_object_id, changed)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_gap_spacing(0, source.drawable_object_id)
            .unwrap(),
        changed
    );
    assert_eq!(
        reopened
            .slide_chart_gap_spacing(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_chart_value_axis_steps_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let defaults = ChartValueAxisSteps::fixed(
        ChartAxisMajorStepCount::new(5).unwrap(),
        ChartAxisMinorStepCount::new(1).unwrap(),
    );
    let fixed = ChartValueAxisSteps::fixed(
        ChartAxisMajorStepCount::new(6).unwrap(),
        ChartAxisMinorStepCount::new(2).unwrap(),
    );
    let major_only = ChartValueAxisSteps::new(Some(ChartAxisMajorStepCount::new(4).unwrap()), None);

    assert_eq!(
        editor
            .slide_chart_value_axis_steps(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_value_axis_steps(0, source.drawable_object_id, defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_value_axis_steps(0, source.drawable_object_id, fixed)
        .unwrap();
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_value_axis_steps(0, duplicate.drawable_object_id)
            .unwrap(),
        fixed
    );

    editor
        .set_slide_chart_value_axis_steps(0, source.drawable_object_id, major_only)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_value_axis_steps(0, source.drawable_object_id)
            .unwrap(),
        major_only
    );

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_value_axis_steps(0, source.drawable_object_id)
            .unwrap(),
        major_only
    );
    assert_eq!(
        reopened
            .slide_chart_value_axis_steps(0, duplicate.drawable_object_id)
            .unwrap(),
        fixed
    );
    reopened
        .set_slide_chart_value_axis_steps(
            0,
            source.drawable_object_id,
            ChartValueAxisSteps::automatic(),
        )
        .unwrap();
    assert_eq!(
        reopened
            .slide_chart_value_axis_steps(0, source.drawable_object_id)
            .unwrap(),
        ChartValueAxisSteps::automatic()
    );
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_presentation_supports_native_chart_value_axis_minimum_label_visibility_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();

    assert!(
        editor
            .slide_chart_value_axis_minimum_label_visible(0, source.drawable_object_id)
            .unwrap()
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_value_axis_minimum_label_visible(0, source.drawable_object_id, true)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_value_axis_minimum_label_visible(0, source.drawable_object_id, false)
        .unwrap();
    assert!(
        !editor
            .slide_chart_value_axis_minimum_label_visible(0, source.drawable_object_id)
            .unwrap()
    );
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert!(
        !editor
            .slide_chart_value_axis_minimum_label_visible(0, duplicate.drawable_object_id)
            .unwrap()
    );

    editor
        .set_slide_chart_value_axis_minimum_label_visible(0, source.drawable_object_id, true)
        .unwrap();
    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert!(
        reopened
            .slide_chart_value_axis_minimum_label_visible(0, source.drawable_object_id)
            .unwrap()
    );
    assert!(
        !reopened
            .slide_chart_value_axis_minimum_label_visible(0, duplicate.drawable_object_id)
            .unwrap()
    );
    reopened
        .set_slide_chart_value_axis_minimum_label_visible(0, source.drawable_object_id, false)
        .unwrap();
    assert!(
        !reopened
            .slide_chart_value_axis_minimum_label_visible(0, source.drawable_object_id)
            .unwrap()
    );
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_presentation_supports_native_chart_category_axis_series_names_visibility_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();

    assert!(
        !editor
            .slide_chart_category_axis_series_names_visible(0, source.drawable_object_id)
            .unwrap()
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_category_axis_series_names_visible(0, source.drawable_object_id, false)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_category_axis_series_names_visible(0, source.drawable_object_id, true)
        .unwrap();
    assert!(
        editor
            .slide_chart_category_axis_series_names_visible(0, source.drawable_object_id)
            .unwrap()
    );
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert!(
        editor
            .slide_chart_category_axis_series_names_visible(0, duplicate.drawable_object_id)
            .unwrap()
    );

    editor
        .set_slide_chart_category_axis_series_names_visible(0, source.drawable_object_id, false)
        .unwrap();
    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert!(
        !reopened
            .slide_chart_category_axis_series_names_visible(0, source.drawable_object_id)
            .unwrap()
    );
    assert!(
        reopened
            .slide_chart_category_axis_series_names_visible(0, duplicate.drawable_object_id)
            .unwrap()
    );
    reopened
        .set_slide_chart_category_axis_series_names_visible(0, source.drawable_object_id, true)
        .unwrap();
    assert!(
        reopened
            .slide_chart_category_axis_series_names_visible(0, source.drawable_object_id)
            .unwrap()
    );
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_presentation_supports_native_chart_axis_label_visibility_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();

    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert!(
            editor
                .slide_chart_axis_labels_visible(0, source.drawable_object_id, axis)
                .unwrap()
        );
    }
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_axis_labels_visible(
            0,
            source.drawable_object_id,
            ChartAxis::Category,
            true,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    for axis in [ChartAxis::Category, ChartAxis::Value] {
        editor
            .set_slide_chart_axis_labels_visible(0, source.drawable_object_id, axis, false)
            .unwrap();
        assert!(
            !editor
                .slide_chart_axis_labels_visible(0, source.drawable_object_id, axis)
                .unwrap()
        );
    }

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert!(
            !editor
                .slide_chart_axis_labels_visible(0, duplicate.drawable_object_id, axis)
                .unwrap()
        );
        editor
            .set_slide_chart_axis_labels_visible(0, source.drawable_object_id, axis, true)
            .unwrap();
    }

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert!(
            reopened
                .slide_chart_axis_labels_visible(0, source.drawable_object_id, axis)
                .unwrap()
        );
        assert!(
            !reopened
                .slide_chart_axis_labels_visible(0, duplicate.drawable_object_id, axis)
                .unwrap()
        );
        reopened
            .set_slide_chart_axis_labels_visible(0, source.drawable_object_id, axis, false)
            .unwrap();
        assert!(
            !reopened
                .slide_chart_axis_labels_visible(0, source.drawable_object_id, axis)
                .unwrap()
        );
    }
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_presentation_supports_native_chart_axis_line_visibility_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();

    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert!(
            editor
                .slide_chart_axis_line_visible(0, source.drawable_object_id, axis)
                .unwrap()
        );
        editor
            .set_slide_chart_axis_line_visible(0, source.drawable_object_id, axis, false)
            .unwrap();
        assert!(
            !editor
                .slide_chart_axis_line_visible(0, source.drawable_object_id, axis)
                .unwrap()
        );
    }

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert!(
            !editor
                .slide_chart_axis_line_visible(0, duplicate.drawable_object_id, axis)
                .unwrap()
        );
        editor
            .set_slide_chart_axis_line_visible(0, source.drawable_object_id, axis, true)
            .unwrap();
    }

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert!(
            reopened
                .slide_chart_axis_line_visible(0, source.drawable_object_id, axis)
                .unwrap()
        );
        assert!(
            !reopened
                .slide_chart_axis_line_visible(0, duplicate.drawable_object_id, axis)
                .unwrap()
        );
    }
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_presentation_supports_native_chart_axis_major_gridline_visibility_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();

    assert!(
        !editor
            .slide_chart_axis_major_gridlines_visible(
                0,
                source.drawable_object_id,
                ChartAxis::Category,
            )
            .unwrap()
    );
    assert!(
        editor
            .slide_chart_axis_major_gridlines_visible(
                0,
                source.drawable_object_id,
                ChartAxis::Value
            )
            .unwrap()
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_axis_major_gridlines_visible(
            0,
            source.drawable_object_id,
            ChartAxis::Category,
            false,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_axis_major_gridlines_visible(
            0,
            source.drawable_object_id,
            ChartAxis::Category,
            true,
        )
        .unwrap();
    editor
        .set_slide_chart_axis_major_gridlines_visible(
            0,
            source.drawable_object_id,
            ChartAxis::Value,
            false,
        )
        .unwrap();

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert_eq!(
            editor
                .slide_chart_axis_major_gridlines_visible(0, duplicate.drawable_object_id, axis)
                .unwrap(),
            axis == ChartAxis::Category
        );
    }

    editor
        .set_slide_chart_axis_major_gridlines_visible(
            0,
            source.drawable_object_id,
            ChartAxis::Category,
            false,
        )
        .unwrap();
    editor
        .set_slide_chart_axis_major_gridlines_visible(
            0,
            source.drawable_object_id,
            ChartAxis::Value,
            true,
        )
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert_eq!(
            reopened
                .slide_chart_axis_major_gridlines_visible(0, source.drawable_object_id, axis)
                .unwrap(),
            axis == ChartAxis::Value
        );
        assert_eq!(
            reopened
                .slide_chart_axis_major_gridlines_visible(0, duplicate.drawable_object_id, axis)
                .unwrap(),
            axis == ChartAxis::Category
        );
    }
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_presentation_supports_native_chart_axis_minor_gridline_visibility_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();

    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert!(
            !editor
                .slide_chart_axis_minor_gridlines_visible(0, source.drawable_object_id, axis)
                .unwrap()
        );
    }
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_axis_minor_gridlines_visible(
            0,
            source.drawable_object_id,
            ChartAxis::Category,
            false,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    for axis in [ChartAxis::Category, ChartAxis::Value] {
        editor
            .set_slide_chart_axis_minor_gridlines_visible(0, source.drawable_object_id, axis, true)
            .unwrap();
    }
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert!(
            editor
                .slide_chart_axis_minor_gridlines_visible(0, duplicate.drawable_object_id, axis)
                .unwrap()
        );
        editor
            .set_slide_chart_axis_minor_gridlines_visible(0, source.drawable_object_id, axis, false)
            .unwrap();
    }

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert!(
            !reopened
                .slide_chart_axis_minor_gridlines_visible(0, source.drawable_object_id, axis)
                .unwrap()
        );
        assert!(
            reopened
                .slide_chart_axis_minor_gridlines_visible(0, duplicate.drawable_object_id, axis)
                .unwrap()
        );
    }
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_presentation_supports_native_chart_axis_minor_tick_mark_visibility_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();

    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert!(
            editor
                .slide_chart_axis_minor_tick_marks_visible(0, source.drawable_object_id, axis)
                .unwrap()
        );
    }
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_axis_minor_tick_marks_visible(
            0,
            source.drawable_object_id,
            ChartAxis::Category,
            true,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    for axis in [ChartAxis::Category, ChartAxis::Value] {
        editor
            .set_slide_chart_axis_minor_tick_marks_visible(
                0,
                source.drawable_object_id,
                axis,
                false,
            )
            .unwrap();
        assert!(
            !editor
                .slide_chart_axis_minor_tick_marks_visible(0, source.drawable_object_id, axis)
                .unwrap()
        );
    }

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert!(
            !editor
                .slide_chart_axis_minor_tick_marks_visible(0, duplicate.drawable_object_id, axis)
                .unwrap()
        );
        editor
            .set_slide_chart_axis_minor_tick_marks_visible(0, source.drawable_object_id, axis, true)
            .unwrap();
    }

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert!(
            reopened
                .slide_chart_axis_minor_tick_marks_visible(0, source.drawable_object_id, axis)
                .unwrap()
        );
        assert!(
            !reopened
                .slide_chart_axis_minor_tick_marks_visible(0, duplicate.drawable_object_id, axis)
                .unwrap()
        );
        reopened
            .set_slide_chart_axis_minor_tick_marks_visible(
                0,
                source.drawable_object_id,
                axis,
                false,
            )
            .unwrap();
        assert!(
            !reopened
                .slide_chart_axis_minor_tick_marks_visible(0, source.drawable_object_id, axis)
                .unwrap()
        );
    }
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_presentation_supports_native_chart_axis_tick_mark_location_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();

    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert_eq!(
            editor
                .slide_chart_axis_tick_mark_location(0, source.drawable_object_id, axis)
                .unwrap(),
            ChartAxisTickMarkLocation::Centered
        );
    }
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_axis_tick_mark_location(
            0,
            source.drawable_object_id,
            ChartAxis::Category,
            ChartAxisTickMarkLocation::Centered,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_axis_tick_mark_location(
            0,
            source.drawable_object_id,
            ChartAxis::Category,
            ChartAxisTickMarkLocation::None,
        )
        .unwrap();
    editor
        .set_slide_chart_axis_tick_mark_location(
            0,
            source.drawable_object_id,
            ChartAxis::Value,
            ChartAxisTickMarkLocation::Outside,
        )
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_axis_tick_mark_location(0, source.drawable_object_id, ChartAxis::Category)
            .unwrap(),
        ChartAxisTickMarkLocation::None
    );
    assert_eq!(
        editor
            .slide_chart_axis_tick_mark_location(0, source.drawable_object_id, ChartAxis::Value)
            .unwrap(),
        ChartAxisTickMarkLocation::Outside
    );

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_axis_tick_mark_location(
                0,
                duplicate.drawable_object_id,
                ChartAxis::Category,
            )
            .unwrap(),
        ChartAxisTickMarkLocation::None
    );
    assert_eq!(
        editor
            .slide_chart_axis_tick_mark_location(0, duplicate.drawable_object_id, ChartAxis::Value)
            .unwrap(),
        ChartAxisTickMarkLocation::Outside
    );

    editor
        .set_slide_chart_axis_tick_mark_location(
            0,
            source.drawable_object_id,
            ChartAxis::Category,
            ChartAxisTickMarkLocation::Inside,
        )
        .unwrap();
    editor
        .set_slide_chart_axis_tick_mark_location(
            0,
            source.drawable_object_id,
            ChartAxis::Value,
            ChartAxisTickMarkLocation::Centered,
        )
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_axis_tick_mark_location(0, source.drawable_object_id, ChartAxis::Category)
            .unwrap(),
        ChartAxisTickMarkLocation::Inside
    );
    assert_eq!(
        reopened
            .slide_chart_axis_tick_mark_location(0, source.drawable_object_id, ChartAxis::Value)
            .unwrap(),
        ChartAxisTickMarkLocation::Centered
    );
    assert_eq!(
        reopened
            .slide_chart_axis_tick_mark_location(
                0,
                duplicate.drawable_object_id,
                ChartAxis::Category,
            )
            .unwrap(),
        ChartAxisTickMarkLocation::None
    );
    assert_eq!(
        reopened
            .slide_chart_axis_tick_mark_location(0, duplicate.drawable_object_id, ChartAxis::Value)
            .unwrap(),
        ChartAxisTickMarkLocation::Outside
    );
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_presentation_supports_native_chart_legend_visibility_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();

    assert!(
        editor
            .slide_chart_legend_visible(0, source.drawable_object_id)
            .unwrap()
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_legend_visible(0, source.drawable_object_id, true)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_slide_chart_legend_visible(0, source.drawable_object_id, false)
        .unwrap();
    assert!(
        !editor
            .slide_chart_legend_visible(0, source.drawable_object_id)
            .unwrap()
    );

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert!(
        !editor
            .slide_chart_legend_visible(0, duplicate.drawable_object_id)
            .unwrap()
    );

    editor
        .set_slide_chart_legend_visible(0, source.drawable_object_id, true)
        .unwrap();
    assert!(
        editor
            .slide_chart_legend_visible(0, source.drawable_object_id)
            .unwrap()
    );
    assert!(
        !editor
            .slide_chart_legend_visible(0, duplicate.drawable_object_id)
            .unwrap()
    );

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert!(
        reopened
            .slide_chart_legend_visible(0, source.drawable_object_id)
            .unwrap()
    );
    assert!(
        !reopened
            .slide_chart_legend_visible(0, duplicate.drawable_object_id)
            .unwrap()
    );
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_exact_chart_legend_fill_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let chart = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let object_id = chart.drawable_object_id;
    let baseline = editor.to_bytes().unwrap();

    assert_eq!(
        editor.slide_chart_legend_fill(0, object_id).unwrap(),
        ChartLegendFill::Inherited
    );
    let solid = ChartLegendFill::Fill(ShapeFill::Solid(
        RgbaColor::new(0.15, 0.35, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
    ));
    editor
        .set_slide_chart_legend_fill(0, object_id, &solid)
        .unwrap();
    assert_eq!(editor.slide_chart_legend_fill(0, object_id).unwrap(), solid);
    assert!(editor.slide_chart_legend_visible(0, object_id).unwrap());

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.slide_chart_legend_fill(0, object_id).unwrap(),
        solid
    );
    reopened
        .set_slide_chart_legend_fill(0, object_id, &ChartLegendFill::Fill(ShapeFill::None))
        .unwrap();
    assert_eq!(
        reopened.slide_chart_legend_fill(0, object_id).unwrap(),
        ChartLegendFill::Fill(ShapeFill::None)
    );
    reopened
        .set_slide_chart_legend_fill(0, object_id, &ChartLegendFill::Inherited)
        .unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), baseline);
}

#[test]
fn scratch_presentation_supports_native_chart_value_axis_scale_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();

    assert_eq!(
        editor
            .slide_chart_value_axis_scale(0, source.drawable_object_id)
            .unwrap(),
        ChartValueAxisScale::Linear
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_value_axis_scale(0, source.drawable_object_id, ChartValueAxisScale::Linear)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_value_axis_scale(
            0,
            source.drawable_object_id,
            ChartValueAxisScale::Logarithmic,
        )
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_value_axis_scale(0, source.drawable_object_id)
            .unwrap(),
        ChartValueAxisScale::Logarithmic
    );

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_value_axis_scale(0, duplicate.drawable_object_id)
            .unwrap(),
        ChartValueAxisScale::Logarithmic
    );
    editor
        .set_slide_chart_value_axis_scale(0, source.drawable_object_id, ChartValueAxisScale::Linear)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_value_axis_scale(0, source.drawable_object_id)
            .unwrap(),
        ChartValueAxisScale::Linear
    );
    assert_eq!(
        reopened
            .slide_chart_value_axis_scale(0, duplicate.drawable_object_id)
            .unwrap(),
        ChartValueAxisScale::Logarithmic
    );
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_presentation_supports_native_chart_border_stroke_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let default = ShapeStroke::new(RgbaColor::black(), StrokeWidth::ONE, StrokePattern::Solid);
    let customized = chart_stroke(StrokePattern::MediumDash, 3.0);
    let changed = chart_stroke(StrokePattern::RoundedDash, 2.0);

    assert_eq!(
        editor
            .slide_chart_border_stroke(0, source.drawable_object_id)
            .unwrap(),
        Some(default)
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_border_stroke(0, source.drawable_object_id, Some(default))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_border_stroke(0, source.drawable_object_id, Some(customized))
        .unwrap();
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_border_stroke(0, duplicate.drawable_object_id)
            .unwrap(),
        Some(customized)
    );
    editor
        .set_slide_chart_border_stroke(0, source.drawable_object_id, Some(changed))
        .unwrap();
    editor
        .set_slide_chart_border_stroke(0, duplicate.drawable_object_id, None)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_border_stroke(0, source.drawable_object_id)
            .unwrap(),
        Some(changed)
    );
    assert_eq!(
        reopened
            .slide_chart_border_stroke(0, duplicate.drawable_object_id)
            .unwrap(),
        None
    );
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_chart_background_fill_crud() {
    let image_bytes = fixture("test-data/images/png/lena.png");
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let native_default = editor
        .slide_chart_background_fill(0, source.drawable_object_id)
        .unwrap();
    assert!(matches!(native_default, ShapeFill::Gradient(_)));
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_background_fill(0, source.drawable_object_id, &native_default)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let customized = chart_background_fill();
    let image = editor
        .set_slide_chart_background_image_fill(
            0,
            source.drawable_object_id,
            "lena.png",
            &image_bytes,
            ShapeImageFillTechnique::Tile,
            None,
        )
        .unwrap();
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_background_fill(0, duplicate.drawable_object_id)
            .unwrap(),
        ShapeFill::Image(image.clone())
    );
    editor
        .set_slide_chart_background_fill(0, source.drawable_object_id, &customized)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_background_fill(0, source.drawable_object_id)
            .unwrap(),
        customized
    );
    assert_eq!(
        reopened
            .slide_chart_background_fill(0, duplicate.drawable_object_id)
            .unwrap(),
        ShapeFill::Image(image.clone())
    );
    assert_eq!(
        reopened
            .extract_media(image.data_identifier().unwrap().get())
            .unwrap(),
        image_bytes
    );
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.media_assets().unwrap().is_empty());
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_inherited_series_fill_crud() {
    let image_bytes = fixture("test-data/images/png/lena.png");
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let defaults = editor
        .slide_chart_series_fills(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(defaults.len(), 2);
    assert!(
        defaults
            .iter()
            .all(|fill| matches!(fill, ShapeFill::Solid(_)))
    );
    assert_ne!(defaults[0], defaults[1]);
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_series_fills(0, source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let first = ChartSeriesIndex::from_zero_based(0);
    let second = ChartSeriesIndex::from_zero_based(1);
    editor
        .set_slide_chart_series_fill(0, source.drawable_object_id, first, &ShapeFill::None)
        .unwrap();
    let image = editor
        .set_slide_chart_series_image_fill(
            0,
            source.drawable_object_id,
            second,
            "lena.png",
            &image_bytes,
            ShapeImageFillTechnique::ScaleToFit,
            None,
        )
        .unwrap();
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_series_fills(0, duplicate.drawable_object_id)
            .unwrap(),
        vec![ShapeFill::None, ShapeFill::Image(image.clone())]
    );
    assert_eq!(
        editor
            .reset_slide_chart_series_fill(0, source.drawable_object_id, first)
            .unwrap(),
        defaults[0]
    );

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_series_fill(0, source.drawable_object_id, first)
            .unwrap(),
        defaults[0]
    );
    assert_eq!(
        reopened
            .slide_chart_series_fill(0, source.drawable_object_id, second)
            .unwrap(),
        ShapeFill::Image(image.clone())
    );
    assert_eq!(
        reopened
            .extract_media(image.data_identifier().unwrap().get())
            .unwrap(),
        image_bytes
    );
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(reopened.media_assets().unwrap().len(), 1);
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.media_assets().unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_inherited_series_stroke_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let defaults = vec![None, None];
    assert_eq!(
        editor
            .slide_chart_series_strokes(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_series_strokes(0, source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let first = ChartSeriesIndex::from_zero_based(0);
    let second = ChartSeriesIndex::from_zero_based(1);
    let rounded = chart_series_stroke(ChartSeriesStrokePattern::RoundedDash, 3.5);
    let medium = chart_series_stroke(ChartSeriesStrokePattern::MediumDash, 2.0);
    editor
        .set_slide_chart_series_strokes(
            0,
            source.drawable_object_id,
            &[Some(rounded), Some(medium)],
        )
        .unwrap();
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_series_strokes(0, duplicate.drawable_object_id)
            .unwrap(),
        vec![Some(rounded), Some(medium)]
    );
    editor
        .set_slide_chart_series_stroke(0, source.drawable_object_id, first, None)
        .unwrap();
    assert_eq!(
        editor
            .reset_slide_chart_series_stroke(0, source.drawable_object_id, first)
            .unwrap(),
        None
    );

    let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_series_stroke(0, source.drawable_object_id, first)
            .unwrap(),
        None
    );
    assert_eq!(
        reopened
            .slide_chart_series_stroke(0, source.drawable_object_id, second)
            .unwrap(),
        Some(medium)
    );
    assert_eq!(
        reopened
            .slide_chart_series_strokes(0, duplicate.drawable_object_id)
            .unwrap(),
        vec![Some(rounded), Some(medium)]
    );
}

#[test]
fn scratch_presentation_supports_native_chart_shadow_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let native_default = ChartShadow::native_default();
    assert_eq!(
        editor
            .slide_chart_shadow(0, source.drawable_object_id)
            .unwrap(),
        native_default
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_shadow(0, source.drawable_object_id, native_default)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let customized = chart_shadow();
    editor
        .set_slide_chart_shadow(0, source.drawable_object_id, customized)
        .unwrap();
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_shadow(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    editor
        .set_slide_chart_shadow(0, source.drawable_object_id, ChartShadow::None)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_shadow(0, source.drawable_object_id)
            .unwrap(),
        ChartShadow::None
    );
    assert_eq!(
        reopened
            .slide_chart_shadow(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    reopened
        .set_slide_chart_shadow(0, duplicate.drawable_object_id, native_default)
        .unwrap();
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_pie_start_angle_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Pie2d, pie_data(), POSITION, SIZE)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_pie_start_angle(0, source.drawable_object_id)
            .unwrap(),
        ChartPieStartAngle::ZERO
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_pie_start_angle(0, source.drawable_object_id, ChartPieStartAngle::ZERO)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let customized = ChartPieStartAngle::from_degrees(123.0).unwrap();
    editor
        .set_slide_chart_pie_start_angle(0, source.drawable_object_id, customized)
        .unwrap();
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    editor
        .set_slide_chart_kind(0, duplicate.drawable_object_id, ChartKind::Donut2d)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_pie_start_angle(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    editor
        .set_slide_chart_pie_start_angle(
            0,
            source.drawable_object_id,
            ChartPieStartAngle::HALF_TURN,
        )
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_pie_start_angle(0, source.drawable_object_id)
            .unwrap(),
        ChartPieStartAngle::HALF_TURN
    );
    assert_eq!(
        reopened
            .slide_chart_pie_start_angle(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    reopened
        .set_slide_chart_pie_start_angle(0, duplicate.drawable_object_id, ChartPieStartAngle::ZERO)
        .unwrap();

    let column = reopened
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let before_rejected_update = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .slide_chart_pie_start_angle(0, column.drawable_object_id)
            .is_err()
    );
    assert!(
        reopened
            .set_slide_chart_pie_start_angle(
                0,
                column.drawable_object_id,
                ChartPieStartAngle::QUARTER_TURN,
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected_update);

    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, column.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_donut_inner_radius_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Donut2d, pie_data(), POSITION, SIZE)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_donut_inner_radius(0, source.drawable_object_id)
            .unwrap(),
        ChartDonutInnerRadius::DEFAULT
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_donut_inner_radius(
            0,
            source.drawable_object_id,
            ChartDonutInnerRadius::DEFAULT,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let customized = ChartDonutInnerRadius::from_percent(42.0).unwrap();
    editor
        .set_slide_chart_donut_inner_radius(0, source.drawable_object_id, customized)
        .unwrap();
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_donut_inner_radius(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    editor
        .set_slide_chart_kind(0, duplicate.drawable_object_id, ChartKind::Pie2d)
        .unwrap();
    let before_rejected_update = editor.to_bytes().unwrap();
    assert!(
        editor
            .slide_chart_donut_inner_radius(0, duplicate.drawable_object_id)
            .is_err()
    );
    assert!(
        editor
            .set_slide_chart_donut_inner_radius(
                0,
                duplicate.drawable_object_id,
                ChartDonutInnerRadius::MAXIMUM,
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_rejected_update);
    editor
        .set_slide_chart_kind(0, duplicate.drawable_object_id, ChartKind::Donut3d)
        .unwrap();

    editor
        .set_slide_chart_donut_inner_radius(
            0,
            source.drawable_object_id,
            ChartDonutInnerRadius::MINIMUM,
        )
        .unwrap();
    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_donut_inner_radius(0, source.drawable_object_id)
            .unwrap(),
        ChartDonutInnerRadius::MINIMUM
    );
    assert_eq!(
        reopened
            .slide_chart_donut_inner_radius(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    reopened
        .set_slide_chart_donut_inner_radius(
            0,
            duplicate.drawable_object_id,
            ChartDonutInnerRadius::DEFAULT,
        )
        .unwrap();
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_pie_wedge_explosion_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Pie2d, pie_data(), POSITION, SIZE)
        .unwrap();
    let zeros = vec![ChartPieWedgeExplosion::ZERO; 3];
    assert_eq!(
        editor
            .slide_chart_pie_wedge_explosions(0, source.drawable_object_id)
            .unwrap(),
        zeros
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_pie_wedge_explosions(0, source.drawable_object_id, &zeros)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let customized = [
        ChartPieWedgeExplosion::from_percent(10.0).unwrap(),
        ChartPieWedgeExplosion::from_percent(25.0).unwrap(),
        ChartPieWedgeExplosion::from_percent(40.0).unwrap(),
    ];
    editor
        .set_slide_chart_pie_wedge_explosions(0, source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_pie_wedge_explosion(
                0,
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(1),
            )
            .unwrap(),
        customized[1]
    );
    editor
        .set_slide_chart_pie_wedge_explosions(0, source.drawable_object_id, &zeros)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_pie_wedge_explosions(0, source.drawable_object_id, &customized)
        .unwrap();
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    editor
        .set_slide_chart_kind(0, duplicate.drawable_object_id, ChartKind::Donut2d)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_pie_wedge_explosions(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    let isolated = ChartPieWedgeExplosion::from_percent(55.0).unwrap();
    editor
        .set_slide_chart_pie_wedge_explosion(
            0,
            source.drawable_object_id,
            ChartPieWedgeIndex::from_zero_based(0),
            isolated,
        )
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_pie_wedge_explosion(
                0,
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(0),
            )
            .unwrap(),
        isolated
    );
    assert_eq!(
        reopened
            .slide_chart_pie_wedge_explosions(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected_updates = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_slide_chart_pie_wedge_explosions(0, source.drawable_object_id, &customized[..2],)
            .is_err()
    );
    assert!(
        reopened
            .set_slide_chart_pie_wedge_explosion(
                0,
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(3),
                isolated,
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected_updates);

    let column = reopened
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let before_wrong_kind = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .slide_chart_pie_wedge_explosions(0, column.drawable_object_id)
            .is_err()
    );
    assert!(
        reopened
            .set_slide_chart_pie_wedge_explosions(0, column.drawable_object_id, &customized,)
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_wrong_kind);

    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, column.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_pie_label_visibility_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Pie2d, pie_data(), POSITION, SIZE)
        .unwrap();
    let defaults = vec![ChartPieLabelVisibility::DEFAULT; 3];
    let customized = [
        ChartPieLabelVisibility::DATA_POINT_NAMES_ONLY,
        ChartPieLabelVisibility::ALL,
        ChartPieLabelVisibility::HIDDEN,
    ];
    assert_eq!(
        editor
            .slide_chart_pie_label_visibilities(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_pie_label_visibilities(0, source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_pie_label_visibilities(0, source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_pie_label_visibility(
                0,
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(1),
            )
            .unwrap(),
        ChartPieLabelVisibility::ALL
    );
    let explosions = [
        ChartPieWedgeExplosion::from_percent(10.0).unwrap(),
        ChartPieWedgeExplosion::from_percent(25.0).unwrap(),
        ChartPieWedgeExplosion::from_percent(40.0).unwrap(),
    ];
    editor
        .set_slide_chart_pie_wedge_explosions(0, source.drawable_object_id, &explosions)
        .unwrap();
    editor
        .set_slide_chart_pie_label_visibilities(0, source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_pie_wedge_explosions(0, source.drawable_object_id)
            .unwrap(),
        explosions
    );
    editor
        .set_slide_chart_pie_wedge_explosions(
            0,
            source.drawable_object_id,
            &[ChartPieWedgeExplosion::ZERO; 3],
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_pie_label_visibilities(0, source.drawable_object_id, &customized)
        .unwrap();
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    editor
        .set_slide_chart_kind(0, duplicate.drawable_object_id, ChartKind::Donut2d)
        .unwrap();
    editor
        .set_slide_chart_pie_label_visibility(
            0,
            source.drawable_object_id,
            ChartPieWedgeIndex::from_zero_based(0),
            ChartPieLabelVisibility::VALUES_ONLY,
        )
        .unwrap();
    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_pie_label_visibilities(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_slide_chart_pie_label_visibilities(0, source.drawable_object_id, &customized[..2],)
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_pie_label_distance_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Pie2d, pie_data(), POSITION, SIZE)
        .unwrap();
    let defaults = vec![ChartPieLabelDistance::DEFAULT; 3];
    let customized = [
        ChartPieLabelDistance::MINIMUM,
        ChartPieLabelDistance::from_percent(100.0).unwrap(),
        ChartPieLabelDistance::MAXIMUM,
    ];
    assert_eq!(
        editor
            .slide_chart_pie_label_distances(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_pie_label_distances(0, source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_pie_label_distances(0, source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_pie_label_distance(
                0,
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(1),
            )
            .unwrap(),
        customized[1]
    );
    let visibilities = [
        ChartPieLabelVisibility::DATA_POINT_NAMES_ONLY,
        ChartPieLabelVisibility::ALL,
        ChartPieLabelVisibility::VALUES_ONLY,
    ];
    editor
        .set_slide_chart_pie_label_visibilities(0, source.drawable_object_id, &visibilities)
        .unwrap();
    editor
        .set_slide_chart_pie_label_distances(0, source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_pie_label_visibilities(0, source.drawable_object_id)
            .unwrap(),
        visibilities
    );
    editor
        .set_slide_chart_pie_label_visibilities(
            0,
            source.drawable_object_id,
            &[ChartPieLabelVisibility::DEFAULT; 3],
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_pie_label_distances(0, source.drawable_object_id, &customized)
        .unwrap();
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    editor
        .set_slide_chart_kind(0, duplicate.drawable_object_id, ChartKind::Donut2d)
        .unwrap();
    editor
        .set_slide_chart_pie_label_distance(
            0,
            source.drawable_object_id,
            ChartPieWedgeIndex::from_zero_based(0),
            ChartPieLabelDistance::DEFAULT,
        )
        .unwrap();
    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_pie_label_distances(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_slide_chart_pie_label_distances(0, source.drawable_object_id, &customized[..2],)
            .is_err()
    );
    assert!(
        reopened
            .slide_chart_pie_label_distance(
                0,
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(3),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_series_value_label_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let defaults = [ChartSeriesValueLabelVisibility::Hidden; 2];
    let customized = [
        ChartSeriesValueLabelVisibility::Visible,
        ChartSeriesValueLabelVisibility::Hidden,
    ];

    assert_eq!(
        editor
            .slide_chart_series_value_label_visibilities(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_series_value_label_visibilities(0, source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_series_value_label_visibilities(0, source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_series_value_label_visibility(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(0),
            )
            .unwrap(),
        ChartSeriesValueLabelVisibility::Visible
    );
    editor
        .set_slide_chart_series_value_label_visibilities(0, source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_series_value_label_visibilities(0, source.drawable_object_id, &customized)
        .unwrap();
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    editor
        .set_slide_chart_series_value_label_visibility(
            0,
            source.drawable_object_id,
            ChartSeriesIndex::from_zero_based(0),
            ChartSeriesValueLabelVisibility::Hidden,
        )
        .unwrap();
    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_series_value_label_visibilities(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .slide_chart_series_value_label_visibilities(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_slide_chart_series_value_label_visibilities(
                0,
                source.drawable_object_id,
                &customized[..1],
            )
            .is_err()
    );
    assert!(
        reopened
            .slide_chart_series_value_label_visibility(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_series_value_label_location_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let defaults = [ChartSeriesValueLabelLocation::Top; 2];
    let customized = [
        ChartSeriesValueLabelLocation::Outside,
        ChartSeriesValueLabelLocation::Top,
    ];

    assert_eq!(
        editor
            .slide_chart_series_value_label_locations(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_series_value_label_locations(0, source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_series_value_label_locations(0, source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_series_value_label_location(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(0),
            )
            .unwrap(),
        ChartSeriesValueLabelLocation::Outside
    );
    editor
        .set_slide_chart_series_value_label_locations(0, source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_series_value_label_locations(0, source.drawable_object_id, &customized)
        .unwrap();
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    editor
        .set_slide_chart_series_value_label_location(
            0,
            source.drawable_object_id,
            ChartSeriesIndex::from_zero_based(0),
            ChartSeriesValueLabelLocation::Top,
        )
        .unwrap();
    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_series_value_label_locations(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .slide_chart_series_value_label_locations(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_slide_chart_series_value_label_locations(
                0,
                source.drawable_object_id,
                &customized[..1],
            )
            .is_err()
    );
    assert!(
        reopened
            .slide_chart_series_value_label_location(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_series_value_label_affix_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let defaults = vec![ChartSeriesValueLabelAffixes::default(); 2];
    let customized = vec![
        ChartSeriesValueLabelAffixes::new("$", " USD"),
        ChartSeriesValueLabelAffixes::new("€", " net"),
    ];

    assert_eq!(
        editor
            .slide_chart_series_value_label_affixes(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_series_value_label_affixes(0, source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_chart_series_value_label_affixes(0, source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_series_value_label_affix(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(0),
            )
            .unwrap()
            .prefix(),
        "$"
    );
    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    editor
        .set_slide_chart_series_value_label_affix(
            0,
            source.drawable_object_id,
            ChartSeriesIndex::from_zero_based(0),
            ChartSeriesValueLabelAffixes::default(),
        )
        .unwrap();
    editor
        .set_slide_chart_series_value_label_affix(
            0,
            source.drawable_object_id,
            ChartSeriesIndex::from_zero_based(1),
            ChartSeriesValueLabelAffixes::default(),
        )
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_series_value_label_affixes(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .slide_chart_series_value_label_affixes(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_slide_chart_series_value_label_affixes(
                0,
                source.drawable_object_id,
                &customized[..1],
            )
            .is_err()
    );
    assert!(
        reopened
            .slide_chart_series_value_label_affix(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_series_value_label_number_format_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let defaults = vec![ChartSeriesValueLabelNumberFormat::NATIVE_DEFAULT; 2];
    let fixed_two = ChartSeriesValueLabelNumberFormat::new(
        ChartSeriesValueLabelDecimalPlaces::fixed(2).unwrap(),
        ChartSeriesValueLabelNegativeStyle::Parentheses,
        false,
    );
    let customized = vec![fixed_two, ChartSeriesValueLabelNumberFormat::NATIVE_DEFAULT];

    assert_eq!(
        editor
            .slide_chart_series_value_label_number_formats(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_series_value_label_number_formats(0, source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_slide_chart_series_value_label_number_formats(
            0,
            source.drawable_object_id,
            &customized,
        )
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_series_value_label_number_format(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(0),
            )
            .unwrap(),
        fixed_two
    );

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    editor
        .set_slide_chart_series_value_label_number_format(
            0,
            source.drawable_object_id,
            ChartSeriesIndex::from_zero_based(0),
            ChartSeriesValueLabelNumberFormat::NATIVE_DEFAULT,
        )
        .unwrap();
    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_series_value_label_number_formats(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .slide_chart_series_value_label_number_formats(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_slide_chart_series_value_label_number_formats(
                0,
                source.drawable_object_id,
                &customized[..1],
            )
            .is_err()
    );
    assert!(
        reopened
            .slide_chart_series_value_label_number_format(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_series_value_label_auto_fit_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let defaults = vec![ChartSeriesValueLabelAutoFit::Enabled; 2];
    let customized = vec![
        ChartSeriesValueLabelAutoFit::Disabled,
        ChartSeriesValueLabelAutoFit::Enabled,
    ];

    assert_eq!(
        editor
            .slide_chart_series_value_label_auto_fits(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_series_value_label_auto_fits(0, source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_slide_chart_series_value_label_auto_fits(0, source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_series_value_label_auto_fit(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(0),
            )
            .unwrap(),
        ChartSeriesValueLabelAutoFit::Disabled
    );

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    editor
        .set_slide_chart_series_value_label_auto_fit(
            0,
            source.drawable_object_id,
            ChartSeriesIndex::from_zero_based(0),
            ChartSeriesValueLabelAutoFit::Enabled,
        )
        .unwrap();
    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_series_value_label_auto_fits(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .slide_chart_series_value_label_auto_fits(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_slide_chart_series_value_label_auto_fits(
                0,
                source.drawable_object_id,
                &customized[..1],
            )
            .is_err()
    );
    assert!(
        reopened
            .slide_chart_series_value_label_auto_fit(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_series_trendline_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let defaults = vec![ChartSeriesTrendline::none(); 2];
    let customized = vec![
        ChartSeriesTrendline::linear()
            .with_legend_name("Revenue fit")
            .unwrap()
            .with_equation_visibility(true)
            .unwrap()
            .with_r_squared_visibility(true)
            .unwrap(),
        ChartSeriesTrendline::moving_average(
            ChartSeriesTrendlineMovingAveragePeriod::new(3).unwrap(),
        )
        .with_legend_visibility(true)
        .unwrap(),
    ];

    assert_eq!(
        editor
            .slide_chart_series_trendlines(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_series_trendlines(0, source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_slide_chart_series_trendlines(0, source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_series_trendline(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(1),
            )
            .unwrap(),
        customized[1]
    );

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    for series in 0..2 {
        editor
            .set_slide_chart_series_trendline(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(series),
                ChartSeriesTrendline::none(),
            )
            .unwrap();
    }
    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_series_trendlines(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .slide_chart_series_trendlines(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_slide_chart_series_trendlines(0, source.drawable_object_id, &customized[..1],)
            .is_err()
    );
    assert!(
        reopened
            .slide_chart_series_trendline(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(2),
            )
            .is_err()
    );
    assert!(ChartSeriesTrendline::unsupported(1).is_err());
    assert!(ChartSeriesTrendlinePolynomialOrder::new(7).is_err());
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}

#[test]
fn scratch_presentation_supports_native_series_error_bar_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let source = editor
        .add_slide_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let defaults = vec![ChartSeriesErrorBars::None; 2];
    let customized = vec![
        ChartSeriesErrorBars::FixedValue {
            direction: ChartErrorBarDirection::PositiveAndNegative,
            value: ChartErrorBarFixedValue::new(12.5).unwrap(),
        },
        ChartSeriesErrorBars::Percentage {
            direction: ChartErrorBarDirection::PositiveOnly,
            percentage: ChartErrorBarPercentage::new(17).unwrap(),
        },
    ];
    let default_auto_fits = vec![ChartSeriesErrorBarAutoFit::Enabled; 2];
    let customized_auto_fits = vec![
        ChartSeriesErrorBarAutoFit::Disabled,
        ChartSeriesErrorBarAutoFit::Enabled,
    ];

    assert_eq!(
        editor
            .slide_chart_series_error_bars(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        editor
            .slide_chart_series_error_bar_auto_fits(0, source.drawable_object_id)
            .unwrap(),
        default_auto_fits
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_slide_chart_series_error_bars(0, source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_slide_chart_series_error_bars(0, source.drawable_object_id, &customized)
        .unwrap();
    editor
        .set_slide_chart_series_error_bar_auto_fits(
            0,
            source.drawable_object_id,
            &customized_auto_fits,
        )
        .unwrap();
    assert_eq!(
        editor
            .slide_chart_series_error_bar(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(1),
            )
            .unwrap(),
        customized[1]
    );
    assert_eq!(
        editor
            .slide_chart_series_error_bar_auto_fit(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(0),
            )
            .unwrap(),
        ChartSeriesErrorBarAutoFit::Disabled
    );

    let duplicate = editor
        .duplicate_slide_chart(0, source.drawable_object_id)
        .unwrap();
    for series in 0..2 {
        editor
            .set_slide_chart_series_error_bar(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(series),
                ChartSeriesErrorBars::None,
            )
            .unwrap();
    }
    editor
        .set_slide_chart_series_error_bar_auto_fits(
            0,
            source.drawable_object_id,
            &default_auto_fits,
        )
        .unwrap();
    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_chart_series_error_bars(0, source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .slide_chart_series_error_bars(0, duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    assert_eq!(
        reopened
            .slide_chart_series_error_bar_auto_fits(0, source.drawable_object_id)
            .unwrap(),
        default_auto_fits
    );
    assert_eq!(
        reopened
            .slide_chart_series_error_bar_auto_fits(0, duplicate.drawable_object_id)
            .unwrap(),
        customized_auto_fits
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_slide_chart_series_error_bars(0, source.drawable_object_id, &customized[..1])
            .is_err()
    );
    assert!(
        reopened
            .set_slide_chart_series_error_bar_auto_fits(
                0,
                source.drawable_object_id,
                &customized_auto_fits[..1],
            )
            .is_err()
    );
    assert!(
        reopened
            .slide_chart_series_error_bar(
                0,
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_slide_chart(0, source.drawable_object_id)
        .unwrap();
    reopened
        .remove_slide_chart(0, duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.slide_charts(0).unwrap().is_empty());
}
