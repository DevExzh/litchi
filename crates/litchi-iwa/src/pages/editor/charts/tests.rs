//! Source-built Pages chart CRUD regression tests.

use std::fs;
use std::path::PathBuf;

use super::*;
use crate::charts::{
    Axis, ChartAxisBound, ChartAxisMajorStepCount, ChartAxisMinorStepCount, ChartCornerRadius,
    ChartDonutInnerRadius, ChartErrorBarDirection, ChartErrorBarFixedValue,
    ChartErrorBarPercentage, ChartFont, ChartFontSize, ChartGapPercentage, ChartGapSpacing,
    ChartLegendFill, ChartLegendFont, ChartLegendFontSize, ChartLegendFrame, ChartLegendRect,
    ChartLegendShadow, ChartLegendStroke, ChartPieLabelDistance, ChartPieLabelVisibility,
    ChartPieLeaderLineVisibility, ChartPieStartAngle, ChartPieWedgeExplosion, ChartPieWedgeIndex,
    ChartRoundedCorners, ChartSeriesErrorBarAutoFit, ChartSeriesErrorBars, ChartSeriesIndex,
    ChartSeriesStroke, ChartSeriesStrokePattern, ChartSeriesTrendline,
    ChartSeriesTrendlineMovingAveragePeriod, ChartSeriesTrendlinePolynomialOrder,
    ChartSeriesValueLabelAffixes, ChartSeriesValueLabelAutoFit, ChartSeriesValueLabelDecimalPlaces,
    ChartSeriesValueLabelLocation, ChartSeriesValueLabelNegativeStyle,
    ChartSeriesValueLabelNumberFormat, ChartSeriesValueLabelVisibility, ChartShadow,
    ChartValueAxisBounds, ChartValueAxisScale, ChartValueAxisSteps, TickMarkLocation,
};
use crate::package_metadata::{
    add_component_external_reference, add_component_object_uuids, component_identifier_for_entry,
    remove_component_external_reference, remove_component_object_uuids,
};
use crate::pages::PagesDocumentBuilder;
use crate::shapes::{
    RgbColorSpace, RgbaColor, ShapeDropShadow, ShapeFill, ShapeImageFillTechnique,
    ShapeShadowAngle, ShapeShadowAppearance, ShapeShadowBlurRadius, ShapeShadowOffset,
    ShapeShadowOpacity, ShapeStroke, StrokePattern, StrokeWidth,
};

const POSITION: DrawablePoint = DrawablePoint { x: 96.0, y: 144.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 360.0,
    height: 240.0,
};

fn sample_data() -> ChartData {
    ChartData::new(
        vec!["North".to_owned(), "South".to_owned()],
        vec!["Q1".to_owned(), "Q2".to_owned(), "Q3".to_owned()],
        vec![
            vec![Some(12.0), Some(18.0), Some(24.0)],
            vec![Some(9.0), Some(21.0), Some(27.0)],
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

fn normalize_private_chart_styles_like_pages(
    editor: &PagesEditor,
    drawable_object_id: u64,
) -> PagesEditor {
    let graph = body_chart_graph(editor, drawable_object_id).unwrap();
    let source_group = graph
        .archive_groups
        .iter()
        .find(|group| !group.style_ids.is_empty())
        .unwrap();
    let style_ids = source_group.style_ids.clone();
    let mut package = editor.package().clone();
    let root = root_document(&package).unwrap();
    let theme_id = root.theme.unwrap().identifier;
    let theme = chart_theme_context(&package, theme_id).unwrap();
    let stylesheet_archive_name = find_object_archive(&package, theme.stylesheet_id).unwrap();
    let stylesheet_component_id =
        component_identifier_for_entry(&package, &stylesheet_archive_name)
            .unwrap()
            .unwrap();

    let mut moved = Vec::with_capacity(style_ids.len());
    package
        .update_archive(&source_group.archive_name, |archive| {
            for identifier in &style_ids {
                moved.push(archive.remove_object(*identifier).unwrap());
            }
            Ok(())
        })
        .unwrap();
    package
        .update_archive(&stylesheet_archive_name, |archive| {
            for object in moved {
                archive.insert_object(object)?;
            }
            let stylesheet = archive.object_mut(theme.stylesheet_id).unwrap();
            for info in &mut stylesheet.archive_info.message_infos {
                info.object_references
                    .retain(|identifier| !style_ids.contains(identifier));
                for field in &mut info.field_infos {
                    field
                        .object_references
                        .retain(|identifier| !style_ids.contains(identifier));
                }
            }
            Ok(())
        })
        .unwrap();

    remove_component_object_uuids(&mut package, source_group.component_id, &style_ids).unwrap();
    add_component_object_uuids(&mut package, stylesheet_component_id, &style_ids).unwrap();
    for identifier in style_ids {
        remove_component_external_reference(
            &mut package,
            stylesheet_component_id,
            source_group.component_id,
            identifier,
        )
        .unwrap();
        add_component_external_reference(
            &mut package,
            source_group.component_id,
            stylesheet_component_id,
            identifier,
        )
        .unwrap();
    }
    PagesEditor::from_package(package).unwrap()
}

#[test]
fn scratch_document_supports_body_chart_crud() {
    let mut editor = PagesEditor::create_with_text("Quarterly results").unwrap();
    let anchor = "Quarterly results".encode_utf16().count();
    assert!(editor.body_charts().unwrap().is_empty());

    let created = editor
        .add_body_chart(anchor, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    assert_eq!(created.anchor_character_index, anchor as u32);
    assert_eq!(created.kind, ChartKind::Column2d);
    assert_eq!(created.direction, ChartSeriesDirection::Rows);
    assert_eq!(created.data, sample_data());
    assert_eq!(editor.body_text().unwrap(), "Quarterly results\u{fffc}");

    let replacement = ChartData::new(
        vec!["Revenue".to_owned()],
        vec!["2026".to_owned(), "2027".to_owned(), "2028".to_owned()],
        vec![vec![Some(30.0), Some(45.0), None]],
    )
    .unwrap();
    editor
        .set_body_chart_kind(created.drawable_object_id, ChartKind::Bar2d)
        .unwrap();
    editor
        .set_body_chart_data(created.drawable_object_id, replacement.clone())
        .unwrap();
    editor
        .set_body_chart_direction(created.drawable_object_id, ChartSeriesDirection::Columns)
        .unwrap();
    let changed_geometry = chart_geometry(
        "Pages",
        DrawablePoint { x: 72.0, y: 216.0 },
        DrawableSize {
            width: 420.0,
            height: 260.0,
        },
    )
    .unwrap();
    editor
        .set_body_chart_geometry(created.drawable_object_id, changed_geometry)
        .unwrap();

    let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let charts = reopened.body_charts().unwrap();
    assert_eq!(charts.len(), 1);
    assert_eq!(charts[0].kind, ChartKind::Bar2d);
    assert_eq!(charts[0].direction, ChartSeriesDirection::Columns);
    assert_eq!(charts[0].data, replacement);
    assert_eq!(charts[0].geometry, changed_geometry);

    let removed = editor
        .remove_body_chart(created.drawable_object_id)
        .unwrap();
    assert_eq!(removed.chart.drawable_object_id, created.drawable_object_id);
    assert_eq!(editor.body_text().unwrap(), "Quarterly results");
    assert!(editor.body_charts().unwrap().is_empty());
    PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
}

#[test]
fn pages_normalized_chart_styles_support_full_lifecycle_crud() {
    let mut editor = PagesEditor::create_with_text("Chart").unwrap();
    let source = editor
        .add_body_chart(5, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let frame = ChartLegendFrame::Frame(ChartLegendRect::from_points(43.0, 8.0, 0.0, 0.0).unwrap());
    editor
        .set_body_chart_legend_frame(source.drawable_object_id, frame)
        .unwrap();
    let mut editor = normalize_private_chart_styles_like_pages(&editor, source.drawable_object_id);

    let charts = editor.body_charts().unwrap();
    assert_eq!(charts.len(), 1);
    assert_eq!(
        editor
            .body_chart_legend_frame(source.drawable_object_id)
            .unwrap(),
        frame
    );
    let changed_frame =
        ChartLegendFrame::Frame(ChartLegendRect::from_points(56.0, 10.0, 0.0, 0.0).unwrap());
    editor
        .set_body_chart_legend_frame(source.drawable_object_id, changed_frame)
        .unwrap();
    let duplicate_anchor = editor.body_text().unwrap().encode_utf16().count();
    let duplicate = editor
        .duplicate_body_chart(source.drawable_object_id, duplicate_anchor)
        .unwrap();
    editor.remove_body_chart(source.drawable_object_id).unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.body_charts().unwrap().len(), 1);
    assert_eq!(
        reopened
            .body_chart_legend_frame(duplicate.drawable_object_id)
            .unwrap(),
        changed_frame
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
    assert_eq!(reopened.body_text().unwrap(), "Chart");
}

#[test]
fn chart_creation_rejects_invalid_inputs_transactionally() {
    let mut editor = PagesEditor::create_with_text("Body").unwrap();
    let baseline = editor.to_bytes().unwrap();
    assert!(
        editor
            .add_body_chart(4, ChartKind::Undefined, sample_data(), POSITION, SIZE)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    assert!(
        editor
            .add_body_chart(5, ChartKind::Column2d, sample_data(), POSITION, SIZE,)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    assert!(
        editor
            .add_body_chart(
                4,
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
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn removing_an_earlier_chart_preserves_later_chart_and_anchor() {
    let mut editor = PagesEditor::create_with_text("Body").unwrap();
    let first = editor
        .add_body_chart(4, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let second = editor
        .add_body_chart(
            5,
            ChartKind::Line2d,
            sample_data(),
            DrawablePoint { x: 120.0, y: 420.0 },
            SIZE,
        )
        .unwrap();

    editor.remove_body_chart(first.drawable_object_id).unwrap();
    let remaining = editor.body_charts().unwrap();
    assert_eq!(remaining.len(), 1);
    assert_eq!(remaining[0].drawable_object_id, second.drawable_object_id);
    assert_eq!(remaining[0].anchor_character_index, 4);
    assert_eq!(editor.body_text().unwrap(), "Body\u{fffc}");

    editor.remove_body_chart(second.drawable_object_id).unwrap();
    assert!(editor.body_charts().unwrap().is_empty());
    assert_eq!(editor.body_text().unwrap(), "Body");
}

#[test]
fn duplicate_body_chart_clones_the_private_graph_and_inline_data() {
    let mut editor = PagesEditor::create_with_text("Body").unwrap();
    let source = editor
        .add_body_chart(4, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let source_graph = body_chart_graph(&editor, source.drawable_object_id).unwrap();
    let baseline = editor.to_bytes().unwrap();
    assert!(editor.duplicate_body_chart(u64::MAX, 5).is_err());
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let duplicate_anchor = editor.body_text().unwrap().encode_utf16().count();
    let duplicate = editor
        .duplicate_body_chart(source.drawable_object_id, duplicate_anchor)
        .unwrap();
    let duplicate_graph = body_chart_graph(&editor, duplicate.drawable_object_id).unwrap();
    let expected_geometry =
        offset_drawable_geometry(source.geometry, BODY_DRAWABLE_DUPLICATE_OFFSET).unwrap();

    assert_ne!(duplicate.drawable_object_id, source.drawable_object_id);
    assert_eq!(duplicate.anchor_character_index, duplicate_anchor as u32);
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
    assert_eq!(editor.body_text().unwrap(), "Body\u{fffc}\u{fffc}");

    let replacement = ChartData::new(
        vec!["Revenue".to_owned()],
        vec!["2026".to_owned(), "2027".to_owned()],
        vec![vec![Some(30.0), Some(45.0)]],
    )
    .unwrap();
    editor
        .set_body_chart_data(duplicate.drawable_object_id, replacement.clone())
        .unwrap();
    assert_eq!(
        body_chart_graph(&editor, source.drawable_object_id)
            .unwrap()
            .info
            .data,
        source.data
    );
    assert_eq!(
        body_chart_graph(&editor, duplicate.drawable_object_id)
            .unwrap()
            .info
            .data,
        replacement
    );

    editor.remove_body_chart(source.drawable_object_id).unwrap();
    assert_eq!(
        editor
            .body_charts()
            .unwrap()
            .iter()
            .map(|chart| chart.drawable_object_id)
            .collect::<Vec<_>>(),
        vec![duplicate.drawable_object_id]
    );
    editor
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(editor.body_charts().unwrap().is_empty());
    assert_eq!(editor.body_text().unwrap(), "Body");
}

#[test]
fn scratch_document_supports_native_chart_caption_crud() {
    let mut editor = PagesEditor::create_with_text("Chart captions").unwrap();
    let source = editor
        .add_body_chart(
            "Chart captions".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    assert_eq!(
        editor
            .body_chart_caption(source.drawable_object_id)
            .unwrap(),
        None
    );
    editor
        .set_body_chart_caption(source.drawable_object_id, "Revenue by region")
        .unwrap();
    assert_eq!(
        editor
            .body_chart_caption(source.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_caption(duplicate.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );

    editor
        .set_body_chart_caption(source.drawable_object_id, "Updated source caption")
        .unwrap();
    assert!(
        editor
            .remove_body_chart_caption(source.drawable_object_id)
            .unwrap()
    );
    assert!(
        !editor
            .remove_body_chart_caption(source.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        editor
            .body_chart_caption(source.drawable_object_id)
            .unwrap(),
        None
    );

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_caption(duplicate.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(
        reopened
            .body_charts()
            .unwrap()
            .iter()
            .all(|chart| chart.drawable_object_id != duplicate.drawable_object_id)
    );
}

#[test]
fn scratch_document_supports_native_chart_title_crud() {
    let mut editor = PagesEditor::create_with_text("Chart titles").unwrap();
    let source = editor
        .add_body_chart(
            "Chart titles".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    assert_eq!(
        editor.body_chart_title(source.drawable_object_id).unwrap(),
        None
    );
    editor
        .set_body_chart_title(source.drawable_object_id, "Revenue by region")
        .unwrap();
    assert_eq!(
        editor.body_chart_title(source.drawable_object_id).unwrap(),
        Some("Revenue by region".to_owned())
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_title(duplicate.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );

    editor
        .set_body_chart_title(source.drawable_object_id, "Updated source title")
        .unwrap();
    assert_eq!(
        editor.body_chart_title(source.drawable_object_id).unwrap(),
        Some("Updated source title".to_owned())
    );
    assert_eq!(
        editor
            .body_chart_title(duplicate.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );
    assert!(
        editor
            .remove_body_chart_title(source.drawable_object_id)
            .unwrap()
    );
    assert!(
        !editor
            .remove_body_chart_title(source.drawable_object_id)
            .unwrap()
    );
    assert_eq!(
        editor.body_chart_title(source.drawable_object_id).unwrap(),
        None
    );

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_title(duplicate.drawable_object_id)
            .unwrap(),
        Some("Revenue by region".to_owned())
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(
        reopened
            .body_charts()
            .unwrap()
            .iter()
            .all(|chart| chart.drawable_object_id != duplicate.drawable_object_id)
    );
}

#[test]
fn scratch_document_supports_native_chart_axis_title_crud() {
    let mut editor = PagesEditor::create_with_text("Chart axis titles").unwrap();
    let source = editor
        .add_body_chart(
            "Chart axis titles".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    for axis in [Axis::Category, Axis::Value] {
        assert_eq!(
            editor
                .body_chart_axis_title(source.drawable_object_id, axis)
                .unwrap(),
            None
        );
    }
    editor
        .set_body_chart_axis_title(source.drawable_object_id, Axis::Category, "Month")
        .unwrap();
    editor
        .set_body_chart_axis_title(source.drawable_object_id, Axis::Value, "Revenue")
        .unwrap();

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for (axis, title) in [(Axis::Category, "Month"), (Axis::Value, "Revenue")] {
        assert_eq!(
            editor
                .body_chart_axis_title(source.drawable_object_id, axis)
                .unwrap()
                .as_deref(),
            Some(title)
        );
        assert_eq!(
            editor
                .body_chart_axis_title(duplicate.drawable_object_id, axis)
                .unwrap()
                .as_deref(),
            Some(title)
        );
    }

    editor
        .set_body_chart_axis_title(source.drawable_object_id, Axis::Category, "Updated month")
        .unwrap();
    editor
        .set_body_chart_axis_title(source.drawable_object_id, Axis::Value, "Updated revenue")
        .unwrap();
    assert_eq!(
        editor
            .body_chart_axis_title(source.drawable_object_id, Axis::Category)
            .unwrap()
            .as_deref(),
        Some("Updated month")
    );
    assert_eq!(
        editor
            .body_chart_axis_title(source.drawable_object_id, Axis::Value)
            .unwrap()
            .as_deref(),
        Some("Updated revenue")
    );
    assert_eq!(
        editor
            .body_chart_axis_title(duplicate.drawable_object_id, Axis::Category)
            .unwrap()
            .as_deref(),
        Some("Month")
    );
    assert_eq!(
        editor
            .body_chart_axis_title(duplicate.drawable_object_id, Axis::Value)
            .unwrap()
            .as_deref(),
        Some("Revenue")
    );

    for axis in [Axis::Category, Axis::Value] {
        assert!(
            editor
                .remove_body_chart_axis_title(source.drawable_object_id, axis)
                .unwrap()
        );
        assert!(
            !editor
                .remove_body_chart_axis_title(source.drawable_object_id, axis)
                .unwrap()
        );
    }

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_axis_title(duplicate.drawable_object_id, Axis::Category)
            .unwrap()
            .as_deref(),
        Some("Month")
    );
    assert_eq!(
        reopened
            .body_chart_axis_title(duplicate.drawable_object_id, Axis::Value)
            .unwrap()
            .as_deref(),
        Some("Revenue")
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(
        reopened
            .body_charts()
            .unwrap()
            .iter()
            .all(|chart| chart.drawable_object_id != duplicate.drawable_object_id)
    );
}

#[test]
fn scratch_document_supports_native_chart_value_axis_bounds_crud() {
    let mut editor = PagesEditor::create_with_text("Chart value-axis bounds").unwrap();
    let source = editor
        .add_body_chart(
            "Chart value-axis bounds".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
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
            .body_chart_value_axis_bounds(source.drawable_object_id)
            .unwrap(),
        automatic
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_value_axis_bounds(source.drawable_object_id, automatic)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_value_axis_bounds(source.drawable_object_id, fixed)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_value_axis_bounds(duplicate.drawable_object_id)
            .unwrap(),
        fixed
    );

    editor
        .set_body_chart_value_axis_bounds(source.drawable_object_id, minimum_only)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_value_axis_bounds(source.drawable_object_id)
            .unwrap(),
        minimum_only
    );

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_value_axis_bounds(source.drawable_object_id)
            .unwrap(),
        minimum_only
    );
    assert_eq!(
        reopened
            .body_chart_value_axis_bounds(duplicate.drawable_object_id)
            .unwrap(),
        fixed
    );
    reopened
        .set_body_chart_value_axis_bounds(source.drawable_object_id, automatic)
        .unwrap();
    assert_eq!(
        reopened
            .body_chart_value_axis_bounds(source.drawable_object_id)
            .unwrap(),
        automatic
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_border_crud() {
    let mut editor = PagesEditor::create_with_text("Chart borders").unwrap();
    let source = editor
        .add_body_chart(
            "Chart borders".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    assert!(
        !editor
            .body_chart_border_visible(source.drawable_object_id)
            .unwrap()
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_border_visible(source.drawable_object_id, false)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_border_visible(source.drawable_object_id, true)
        .unwrap();
    assert!(
        editor
            .body_chart_border_visible(source.drawable_object_id)
            .unwrap()
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert!(
        editor
            .body_chart_border_visible(duplicate.drawable_object_id)
            .unwrap()
    );
    editor
        .set_body_chart_border_visible(source.drawable_object_id, false)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert!(
        !reopened
            .body_chart_border_visible(source.drawable_object_id)
            .unwrap()
    );
    assert!(
        reopened
            .body_chart_border_visible(duplicate.drawable_object_id)
            .unwrap()
    );
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_chart_rounded_corner_crud() {
    let mut editor = PagesEditor::create_with_text("Rounded chart corners").unwrap();
    let source = editor
        .add_body_chart(
            "Rounded chart corners".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let rounded = ChartRoundedCorners::new(ChartCornerRadius::new(20.0).unwrap(), true);
    let changed = ChartRoundedCorners::new(ChartCornerRadius::new(35.0).unwrap(), false);

    assert_eq!(
        editor
            .body_chart_rounded_corners(source.drawable_object_id)
            .unwrap(),
        ChartRoundedCorners::NONE
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_rounded_corners(source.drawable_object_id, ChartRoundedCorners::NONE)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_rounded_corners(source.drawable_object_id, rounded)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_rounded_corners(source.drawable_object_id)
            .unwrap(),
        rounded
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_rounded_corners(duplicate.drawable_object_id)
            .unwrap(),
        rounded
    );
    editor
        .set_body_chart_rounded_corners(source.drawable_object_id, changed)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_rounded_corners(source.drawable_object_id)
            .unwrap(),
        changed
    );
    assert_eq!(
        reopened
            .body_chart_rounded_corners(duplicate.drawable_object_id)
            .unwrap(),
        rounded
    );
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_chart_gap_crud() {
    let mut editor = PagesEditor::create_with_text("Chart gaps").unwrap();
    let source = editor
        .add_body_chart(
            "Chart gaps".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let customized = gap_spacing(25.0, 70.0);
    let changed = gap_spacing(30.0, 60.0);

    assert_eq!(
        editor
            .body_chart_gap_spacing(source.drawable_object_id)
            .unwrap(),
        ChartGapSpacing::NATIVE_DEFAULT
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_gap_spacing(source.drawable_object_id, ChartGapSpacing::NATIVE_DEFAULT)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_gap_spacing(source.drawable_object_id, customized)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_gap_spacing(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    editor
        .set_body_chart_gap_spacing(source.drawable_object_id, changed)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_gap_spacing(source.drawable_object_id)
            .unwrap(),
        changed
    );
    assert_eq!(
        reopened
            .body_chart_gap_spacing(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_chart_value_axis_steps_crud() {
    let mut editor = PagesEditor::create_with_text("Chart value-axis steps").unwrap();
    let source = editor
        .add_body_chart(
            "Chart value-axis steps".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
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
            .body_chart_value_axis_steps(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_value_axis_steps(source.drawable_object_id, defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_value_axis_steps(source.drawable_object_id, fixed)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_value_axis_steps(duplicate.drawable_object_id)
            .unwrap(),
        fixed
    );

    editor
        .set_body_chart_value_axis_steps(source.drawable_object_id, major_only)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_value_axis_steps(source.drawable_object_id)
            .unwrap(),
        major_only
    );

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_value_axis_steps(source.drawable_object_id)
            .unwrap(),
        major_only
    );
    assert_eq!(
        reopened
            .body_chart_value_axis_steps(duplicate.drawable_object_id)
            .unwrap(),
        fixed
    );
    reopened
        .set_body_chart_value_axis_steps(
            source.drawable_object_id,
            ChartValueAxisSteps::automatic(),
        )
        .unwrap();
    assert_eq!(
        reopened
            .body_chart_value_axis_steps(source.drawable_object_id)
            .unwrap(),
        ChartValueAxisSteps::automatic()
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_value_axis_minimum_label_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Chart minimum label").unwrap();
    let source = editor
        .add_body_chart(
            "Chart minimum label".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    assert!(
        editor
            .body_chart_value_axis_minimum_label_visible(source.drawable_object_id)
            .unwrap()
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_value_axis_minimum_label_visible(source.drawable_object_id, true)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_value_axis_minimum_label_visible(source.drawable_object_id, false)
        .unwrap();
    assert!(
        !editor
            .body_chart_value_axis_minimum_label_visible(source.drawable_object_id)
            .unwrap()
    );
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert!(
        !editor
            .body_chart_value_axis_minimum_label_visible(duplicate.drawable_object_id)
            .unwrap()
    );

    editor
        .set_body_chart_value_axis_minimum_label_visible(source.drawable_object_id, true)
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert!(
        reopened
            .body_chart_value_axis_minimum_label_visible(source.drawable_object_id)
            .unwrap()
    );
    assert!(
        !reopened
            .body_chart_value_axis_minimum_label_visible(duplicate.drawable_object_id)
            .unwrap()
    );
    reopened
        .set_body_chart_value_axis_minimum_label_visible(source.drawable_object_id, false)
        .unwrap();
    assert!(
        !reopened
            .body_chart_value_axis_minimum_label_visible(source.drawable_object_id)
            .unwrap()
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_category_axis_series_names_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Chart series names").unwrap();
    let source = editor
        .add_body_chart(
            "Chart series names".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    assert!(
        !editor
            .body_chart_category_axis_series_names_visible(source.drawable_object_id)
            .unwrap()
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_category_axis_series_names_visible(source.drawable_object_id, false)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_category_axis_series_names_visible(source.drawable_object_id, true)
        .unwrap();
    assert!(
        editor
            .body_chart_category_axis_series_names_visible(source.drawable_object_id)
            .unwrap()
    );
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert!(
        editor
            .body_chart_category_axis_series_names_visible(duplicate.drawable_object_id)
            .unwrap()
    );

    editor
        .set_body_chart_category_axis_series_names_visible(source.drawable_object_id, false)
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert!(
        !reopened
            .body_chart_category_axis_series_names_visible(source.drawable_object_id)
            .unwrap()
    );
    assert!(
        reopened
            .body_chart_category_axis_series_names_visible(duplicate.drawable_object_id)
            .unwrap()
    );
    reopened
        .set_body_chart_category_axis_series_names_visible(source.drawable_object_id, true)
        .unwrap();
    assert!(
        reopened
            .body_chart_category_axis_series_names_visible(source.drawable_object_id)
            .unwrap()
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_axis_label_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Chart axis labels").unwrap();
    let source = editor
        .add_body_chart(
            "Chart axis labels".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    for axis in [Axis::Category, Axis::Value] {
        assert!(
            editor
                .body_chart_axis_labels_visible(source.drawable_object_id, axis)
                .unwrap()
        );
    }
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_axis_labels_visible(source.drawable_object_id, Axis::Category, true)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    for axis in [Axis::Category, Axis::Value] {
        editor
            .set_body_chart_axis_labels_visible(source.drawable_object_id, axis, false)
            .unwrap();
        assert!(
            !editor
                .body_chart_axis_labels_visible(source.drawable_object_id, axis)
                .unwrap()
        );
    }

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert!(
            !editor
                .body_chart_axis_labels_visible(duplicate.drawable_object_id, axis)
                .unwrap()
        );
        editor
            .set_body_chart_axis_labels_visible(source.drawable_object_id, axis, true)
            .unwrap();
    }

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert!(
            reopened
                .body_chart_axis_labels_visible(source.drawable_object_id, axis)
                .unwrap()
        );
        assert!(
            !reopened
                .body_chart_axis_labels_visible(duplicate.drawable_object_id, axis)
                .unwrap()
        );
        reopened
            .set_body_chart_axis_labels_visible(source.drawable_object_id, axis, false)
            .unwrap();
        assert!(
            !reopened
                .body_chart_axis_labels_visible(source.drawable_object_id, axis)
                .unwrap()
        );
    }
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_axis_line_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Chart axis lines").unwrap();
    let source = editor
        .add_body_chart(
            "Chart axis lines".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    for axis in [Axis::Category, Axis::Value] {
        assert!(
            editor
                .body_chart_axis_line_visible(source.drawable_object_id, axis)
                .unwrap()
        );
        editor
            .set_body_chart_axis_line_visible(source.drawable_object_id, axis, false)
            .unwrap();
        assert!(
            !editor
                .body_chart_axis_line_visible(source.drawable_object_id, axis)
                .unwrap()
        );
    }

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert!(
            !editor
                .body_chart_axis_line_visible(duplicate.drawable_object_id, axis)
                .unwrap()
        );
        editor
            .set_body_chart_axis_line_visible(source.drawable_object_id, axis, true)
            .unwrap();
    }

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert!(
            reopened
                .body_chart_axis_line_visible(source.drawable_object_id, axis)
                .unwrap()
        );
        assert!(
            !reopened
                .body_chart_axis_line_visible(duplicate.drawable_object_id, axis)
                .unwrap()
        );
    }
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_axis_major_gridline_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Chart major gridlines").unwrap();
    let source = editor
        .add_body_chart(
            "Chart major gridlines".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    assert!(
        !editor
            .body_chart_axis_major_gridlines_visible(source.drawable_object_id, Axis::Category)
            .unwrap()
    );
    assert!(
        editor
            .body_chart_axis_major_gridlines_visible(source.drawable_object_id, Axis::Value)
            .unwrap()
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_axis_major_gridlines_visible(
            source.drawable_object_id,
            Axis::Category,
            false,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_axis_major_gridlines_visible(
            source.drawable_object_id,
            Axis::Category,
            true,
        )
        .unwrap();
    editor
        .set_body_chart_axis_major_gridlines_visible(source.drawable_object_id, Axis::Value, false)
        .unwrap();

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert_eq!(
            editor
                .body_chart_axis_major_gridlines_visible(duplicate.drawable_object_id, axis)
                .unwrap(),
            axis == Axis::Category
        );
    }

    editor
        .set_body_chart_axis_major_gridlines_visible(
            source.drawable_object_id,
            Axis::Category,
            false,
        )
        .unwrap();
    editor
        .set_body_chart_axis_major_gridlines_visible(source.drawable_object_id, Axis::Value, true)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert_eq!(
            reopened
                .body_chart_axis_major_gridlines_visible(source.drawable_object_id, axis)
                .unwrap(),
            axis == Axis::Value
        );
        assert_eq!(
            reopened
                .body_chart_axis_major_gridlines_visible(duplicate.drawable_object_id, axis)
                .unwrap(),
            axis == Axis::Category
        );
    }
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_axis_minor_gridline_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Chart minor gridlines").unwrap();
    let source = editor
        .add_body_chart(
            "Chart minor gridlines".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    for axis in [Axis::Category, Axis::Value] {
        assert!(
            !editor
                .body_chart_axis_minor_gridlines_visible(source.drawable_object_id, axis)
                .unwrap()
        );
    }
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_axis_minor_gridlines_visible(
            source.drawable_object_id,
            Axis::Category,
            false,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    for axis in [Axis::Category, Axis::Value] {
        editor
            .set_body_chart_axis_minor_gridlines_visible(source.drawable_object_id, axis, true)
            .unwrap();
    }
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert!(
            editor
                .body_chart_axis_minor_gridlines_visible(duplicate.drawable_object_id, axis)
                .unwrap()
        );
        editor
            .set_body_chart_axis_minor_gridlines_visible(source.drawable_object_id, axis, false)
            .unwrap();
    }

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert!(
            !reopened
                .body_chart_axis_minor_gridlines_visible(source.drawable_object_id, axis)
                .unwrap()
        );
        assert!(
            reopened
                .body_chart_axis_minor_gridlines_visible(duplicate.drawable_object_id, axis)
                .unwrap()
        );
    }
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_axis_minor_tick_mark_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Chart minor tick marks").unwrap();
    let source = editor
        .add_body_chart(
            "Chart minor tick marks".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    for axis in [Axis::Category, Axis::Value] {
        assert!(
            editor
                .body_chart_axis_minor_tick_marks_visible(source.drawable_object_id, axis)
                .unwrap()
        );
    }
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_axis_minor_tick_marks_visible(
            source.drawable_object_id,
            Axis::Category,
            true,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    for axis in [Axis::Category, Axis::Value] {
        editor
            .set_body_chart_axis_minor_tick_marks_visible(source.drawable_object_id, axis, false)
            .unwrap();
        assert!(
            !editor
                .body_chart_axis_minor_tick_marks_visible(source.drawable_object_id, axis)
                .unwrap()
        );
    }

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert!(
            !editor
                .body_chart_axis_minor_tick_marks_visible(duplicate.drawable_object_id, axis)
                .unwrap()
        );
        editor
            .set_body_chart_axis_minor_tick_marks_visible(source.drawable_object_id, axis, true)
            .unwrap();
    }

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    for axis in [Axis::Category, Axis::Value] {
        assert!(
            reopened
                .body_chart_axis_minor_tick_marks_visible(source.drawable_object_id, axis)
                .unwrap()
        );
        assert!(
            !reopened
                .body_chart_axis_minor_tick_marks_visible(duplicate.drawable_object_id, axis)
                .unwrap()
        );
        reopened
            .set_body_chart_axis_minor_tick_marks_visible(source.drawable_object_id, axis, false)
            .unwrap();
        assert!(
            !reopened
                .body_chart_axis_minor_tick_marks_visible(source.drawable_object_id, axis)
                .unwrap()
        );
    }
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_axis_tick_mark_location_crud() {
    let mut editor = PagesEditor::create_with_text("Chart tick-mark locations").unwrap();
    let source = editor
        .add_body_chart(
            "Chart tick-mark locations".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    for axis in [Axis::Category, Axis::Value] {
        assert_eq!(
            editor
                .body_chart_axis_tick_mark_location(source.drawable_object_id, axis)
                .unwrap(),
            TickMarkLocation::Centered
        );
    }
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_axis_tick_mark_location(
            source.drawable_object_id,
            Axis::Category,
            TickMarkLocation::Centered,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_axis_tick_mark_location(
            source.drawable_object_id,
            Axis::Category,
            TickMarkLocation::None,
        )
        .unwrap();
    editor
        .set_body_chart_axis_tick_mark_location(
            source.drawable_object_id,
            Axis::Value,
            TickMarkLocation::Outside,
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_axis_tick_mark_location(source.drawable_object_id, Axis::Category)
            .unwrap(),
        TickMarkLocation::None
    );
    assert_eq!(
        editor
            .body_chart_axis_tick_mark_location(source.drawable_object_id, Axis::Value)
            .unwrap(),
        TickMarkLocation::Outside
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_axis_tick_mark_location(duplicate.drawable_object_id, Axis::Category)
            .unwrap(),
        TickMarkLocation::None
    );
    assert_eq!(
        editor
            .body_chart_axis_tick_mark_location(duplicate.drawable_object_id, Axis::Value)
            .unwrap(),
        TickMarkLocation::Outside
    );

    editor
        .set_body_chart_axis_tick_mark_location(
            source.drawable_object_id,
            Axis::Category,
            TickMarkLocation::Inside,
        )
        .unwrap();
    editor
        .set_body_chart_axis_tick_mark_location(
            source.drawable_object_id,
            Axis::Value,
            TickMarkLocation::Centered,
        )
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_axis_tick_mark_location(source.drawable_object_id, Axis::Category)
            .unwrap(),
        TickMarkLocation::Inside
    );
    assert_eq!(
        reopened
            .body_chart_axis_tick_mark_location(source.drawable_object_id, Axis::Value)
            .unwrap(),
        TickMarkLocation::Centered
    );
    assert_eq!(
        reopened
            .body_chart_axis_tick_mark_location(duplicate.drawable_object_id, Axis::Category)
            .unwrap(),
        TickMarkLocation::None
    );
    assert_eq!(
        reopened
            .body_chart_axis_tick_mark_location(duplicate.drawable_object_id, Axis::Value)
            .unwrap(),
        TickMarkLocation::Outside
    );
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_legend_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Chart legends").unwrap();
    let source = editor
        .add_body_chart(
            "Chart legends".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    assert!(
        editor
            .body_chart_legend_visible(source.drawable_object_id)
            .unwrap()
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_legend_visible(source.drawable_object_id, true)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_body_chart_legend_visible(source.drawable_object_id, false)
        .unwrap();
    assert!(
        !editor
            .body_chart_legend_visible(source.drawable_object_id)
            .unwrap()
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert!(
        !editor
            .body_chart_legend_visible(duplicate.drawable_object_id)
            .unwrap()
    );

    editor
        .set_body_chart_legend_visible(source.drawable_object_id, true)
        .unwrap();
    assert!(
        editor
            .body_chart_legend_visible(source.drawable_object_id)
            .unwrap()
    );
    assert!(
        !editor
            .body_chart_legend_visible(duplicate.drawable_object_id)
            .unwrap()
    );

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert!(
        reopened
            .body_chart_legend_visible(source.drawable_object_id)
            .unwrap()
    );
    assert!(
        !reopened
            .body_chart_legend_visible(duplicate.drawable_object_id)
            .unwrap()
    );
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_exact_chart_legend_fill_crud() {
    let mut editor = PagesEditor::create_with_text("Chart legend fill").unwrap();
    let chart = editor
        .add_body_chart(
            "Chart legend fill".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let object_id = chart.drawable_object_id;
    let baseline = editor.to_bytes().unwrap();

    assert_eq!(
        editor.body_chart_legend_fill(object_id).unwrap(),
        ChartLegendFill::Inherited
    );
    let solid = ChartLegendFill::Fill(ShapeFill::Solid(
        RgbaColor::new(0.85, 0.25, 0.2, 1.0, RgbColorSpace::Srgb).unwrap(),
    ));
    editor
        .set_body_chart_legend_fill(object_id, &solid)
        .unwrap();
    assert_eq!(editor.body_chart_legend_fill(object_id).unwrap(), solid);
    assert!(editor.body_chart_legend_visible(object_id).unwrap());

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.body_chart_legend_fill(object_id).unwrap(), solid);
    reopened
        .set_body_chart_legend_fill(object_id, &ChartLegendFill::Inherited)
        .unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), baseline);
}

#[test]
fn scratch_document_supports_exact_chart_legend_frame_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Legend")
        .build()
        .unwrap();
    let chart = editor
        .add_body_chart(
            "Legend".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let object_id = chart.drawable_object_id;
    let baseline = editor.to_bytes().unwrap();

    assert_eq!(
        editor.body_chart_legend_frame(object_id).unwrap(),
        ChartLegendFrame::Automatic
    );
    let frame =
        ChartLegendFrame::Frame(ChartLegendRect::from_points(36.0, 18.0, 0.0, 0.0).unwrap());
    editor
        .set_body_chart_legend_frame(object_id, frame)
        .unwrap();
    assert_eq!(editor.body_chart_legend_frame(object_id).unwrap(), frame);

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.body_chart_legend_frame(object_id).unwrap(), frame);
    reopened
        .set_body_chart_legend_frame(object_id, ChartLegendFrame::Automatic)
        .unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), baseline);
}

#[test]
fn scratch_document_supports_exact_chart_legend_typography_crud() {
    let mut editor = PagesDocumentBuilder::new().build().unwrap();
    let chart = editor
        .add_body_chart(0, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let object_id = chart.drawable_object_id;
    let baseline = editor.to_bytes().unwrap();

    assert_eq!(
        editor.body_chart_legend_font(object_id).unwrap(),
        ChartLegendFont::Inherited
    );
    let bold = ChartLegendFont::Font(ChartFont::named("AvenirNext-Bold").unwrap().with_bold(true));
    editor.set_body_chart_legend_font(object_id, &bold).unwrap();
    assert_eq!(editor.body_chart_legend_font(object_id).unwrap(), bold);

    assert_eq!(
        editor.body_chart_legend_font_size(object_id).unwrap(),
        ChartLegendFontSize::Inherited
    );
    let eighteen = ChartLegendFontSize::Size(ChartFontSize::from_points(18.0).unwrap());
    editor
        .set_body_chart_legend_font_size(object_id, eighteen)
        .unwrap();
    assert_eq!(
        editor.body_chart_legend_font_size(object_id).unwrap(),
        eighteen
    );

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.body_chart_legend_font(object_id).unwrap(), bold);
    let italic = ChartLegendFont::Font(
        ChartFont::named("AvenirNext-Italic")
            .unwrap()
            .with_italic(true),
    );
    reopened
        .set_body_chart_legend_font(object_id, &italic)
        .unwrap();
    let fifteen = ChartLegendFontSize::Size(ChartFontSize::from_points(15.0).unwrap());
    reopened
        .set_body_chart_legend_font_size(object_id, fifteen)
        .unwrap();
    assert_eq!(
        reopened.body_chart_legend_font_size(object_id).unwrap(),
        fifteen
    );
    reopened
        .set_body_chart_legend_font(object_id, &ChartLegendFont::Inherited)
        .unwrap();
    assert_eq!(
        reopened.body_chart_legend_font_size(object_id).unwrap(),
        fifteen
    );
    reopened
        .set_body_chart_legend_font_size(object_id, ChartLegendFontSize::Inherited)
        .unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), baseline);
}

#[test]
fn scratch_document_supports_exact_chart_legend_stroke_crud() {
    let mut editor = PagesEditor::create_with_text("Chart legend stroke").unwrap();
    let chart = editor
        .add_body_chart(
            "Chart legend stroke".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let object_id = chart.drawable_object_id;
    let baseline = editor.to_bytes().unwrap();

    assert_eq!(
        editor.body_chart_legend_stroke(object_id).unwrap(),
        ChartLegendStroke::Inherited
    );
    let stroke = ChartLegendStroke::Stroke(ShapeStroke::new(
        RgbaColor::new(0.8, 0.2, 0.15, 1.0, RgbColorSpace::Srgb).unwrap(),
        StrokeWidth::new(1.5).unwrap(),
        StrokePattern::RoundedDash,
    ));
    editor
        .set_body_chart_legend_stroke(object_id, stroke)
        .unwrap();
    assert_eq!(editor.body_chart_legend_stroke(object_id).unwrap(), stroke);
    assert!(editor.body_chart_legend_visible(object_id).unwrap());

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.body_chart_legend_stroke(object_id).unwrap(),
        stroke
    );
    reopened
        .set_body_chart_legend_stroke(object_id, ChartLegendStroke::Inherited)
        .unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), baseline);
}

#[test]
fn scratch_document_supports_exact_chart_legend_shadow_crud() {
    let mut editor = PagesEditor::create_with_text("Chart legend shadow").unwrap();
    let chart = editor
        .add_body_chart(
            "Chart legend shadow".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let object_id = chart.drawable_object_id;
    let baseline = editor.to_bytes().unwrap();

    assert_eq!(
        editor.body_chart_legend_shadow(object_id).unwrap(),
        ChartLegendShadow::Inherited
    );
    let shadow = ChartLegendShadow::Shadow(ShapeDropShadow::new(
        ShapeShadowAppearance::new(
            RgbaColor::black(),
            ShapeShadowBlurRadius::from_points(11).unwrap(),
            ShapeShadowOffset::from_points(7.0).unwrap(),
            ShapeShadowOpacity::new(0.55).unwrap(),
        ),
        ShapeShadowAngle::from_degrees(25.0).unwrap(),
    ));
    editor
        .set_body_chart_legend_shadow(object_id, shadow)
        .unwrap();
    assert_eq!(editor.body_chart_legend_shadow(object_id).unwrap(), shadow);
    assert!(editor.body_chart_legend_visible(object_id).unwrap());

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.body_chart_legend_shadow(object_id).unwrap(),
        shadow
    );
    reopened
        .set_body_chart_legend_shadow(object_id, ChartLegendShadow::Inherited)
        .unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), baseline);
}

#[test]
fn scratch_document_supports_native_chart_value_axis_scale_crud() {
    let mut editor = PagesEditor::create_with_text("Chart value-axis scale").unwrap();
    let source = editor
        .add_body_chart(
            "Chart value-axis scale".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();

    assert_eq!(
        editor
            .body_chart_value_axis_scale(source.drawable_object_id)
            .unwrap(),
        ChartValueAxisScale::Linear
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_value_axis_scale(source.drawable_object_id, ChartValueAxisScale::Linear)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_value_axis_scale(
            source.drawable_object_id,
            ChartValueAxisScale::Logarithmic,
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_value_axis_scale(source.drawable_object_id)
            .unwrap(),
        ChartValueAxisScale::Logarithmic
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_value_axis_scale(duplicate.drawable_object_id)
            .unwrap(),
        ChartValueAxisScale::Logarithmic
    );
    editor
        .set_body_chart_value_axis_scale(source.drawable_object_id, ChartValueAxisScale::Linear)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_value_axis_scale(source.drawable_object_id)
            .unwrap(),
        ChartValueAxisScale::Linear
    );
    assert_eq!(
        reopened
            .body_chart_value_axis_scale(duplicate.drawable_object_id)
            .unwrap(),
        ChartValueAxisScale::Logarithmic
    );
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
}

#[test]
fn scratch_document_supports_native_chart_border_stroke_crud() {
    let mut editor = PagesEditor::create_with_text("Chart border stroke").unwrap();
    let source = editor
        .add_body_chart(
            "Chart border stroke".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let default = ShapeStroke::new(RgbaColor::black(), StrokeWidth::ONE, StrokePattern::Solid);
    let customized = chart_stroke(StrokePattern::MediumDash, 3.0);
    let changed = chart_stroke(StrokePattern::RoundedDash, 2.0);

    assert_eq!(
        editor
            .body_chart_border_stroke(source.drawable_object_id)
            .unwrap(),
        Some(default)
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_border_stroke(source.drawable_object_id, Some(default))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_border_stroke(source.drawable_object_id, Some(customized))
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_border_stroke(duplicate.drawable_object_id)
            .unwrap(),
        Some(customized)
    );
    editor
        .set_body_chart_border_stroke(source.drawable_object_id, Some(changed))
        .unwrap();
    editor
        .set_body_chart_border_stroke(duplicate.drawable_object_id, None)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_border_stroke(source.drawable_object_id)
            .unwrap(),
        Some(changed)
    );
    assert_eq!(
        reopened
            .body_chart_border_stroke(duplicate.drawable_object_id)
            .unwrap(),
        None
    );
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_chart_background_fill_crud() {
    let image_bytes = fixture("test-data/images/png/lena.png");
    let mut editor = PagesEditor::create_with_text("Chart background fill").unwrap();
    let source = editor
        .add_body_chart(
            "Chart background fill".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let native_default = editor
        .body_chart_background_fill(source.drawable_object_id)
        .unwrap();
    assert!(matches!(native_default, ShapeFill::Gradient(_)));
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_background_fill(source.drawable_object_id, &native_default)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let customized = chart_background_fill();
    let image = editor
        .set_body_chart_background_image_fill(
            source.drawable_object_id,
            "lena.png",
            &image_bytes,
            ShapeImageFillTechnique::ScaleToFill,
            None,
        )
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_background_fill(duplicate.drawable_object_id)
            .unwrap(),
        ShapeFill::Image(image.clone())
    );
    editor
        .set_body_chart_background_fill(source.drawable_object_id, &customized)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_background_fill(source.drawable_object_id)
            .unwrap(),
        customized
    );
    assert_eq!(
        reopened
            .body_chart_background_fill(duplicate.drawable_object_id)
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
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.media_assets().unwrap().is_empty());
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_inherited_series_fill_crud() {
    let image_bytes = fixture("test-data/images/png/lena.png");
    let mut editor = PagesEditor::create_with_text("Chart").unwrap();
    let source = editor
        .add_body_chart(5, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let defaults = editor
        .body_chart_series_fills(source.drawable_object_id)
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
        .set_body_chart_series_fills(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let first = ChartSeriesIndex::from_zero_based(0);
    let second = ChartSeriesIndex::from_zero_based(1);
    editor
        .set_body_chart_series_fill(source.drawable_object_id, first, &ShapeFill::None)
        .unwrap();
    let image = editor
        .set_body_chart_series_image_fill(
            source.drawable_object_id,
            second,
            "lena.png",
            &image_bytes,
            ShapeImageFillTechnique::ScaleToFit,
            None,
        )
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(source.drawable_object_id, 6)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_fills(duplicate.drawable_object_id)
            .unwrap(),
        vec![ShapeFill::None, ShapeFill::Image(image.clone())]
    );
    assert_eq!(
        editor
            .reset_body_chart_series_fill(source.drawable_object_id, first)
            .unwrap(),
        defaults[0]
    );

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_fill(source.drawable_object_id, first)
            .unwrap(),
        defaults[0]
    );
    assert_eq!(
        reopened
            .body_chart_series_fill(source.drawable_object_id, second)
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
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    assert_eq!(reopened.media_assets().unwrap().len(), 1);
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.media_assets().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_inherited_series_stroke_crud() {
    let mut editor = PagesEditor::create_with_text("Chart").unwrap();
    let source = editor
        .add_body_chart(5, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let defaults = vec![None, None];
    assert_eq!(
        editor
            .body_chart_series_strokes(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_strokes(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let first = ChartSeriesIndex::from_zero_based(0);
    let second = ChartSeriesIndex::from_zero_based(1);
    let rounded = chart_series_stroke(ChartSeriesStrokePattern::RoundedDash, 3.5);
    let medium = chart_series_stroke(ChartSeriesStrokePattern::MediumDash, 2.0);
    editor
        .set_body_chart_series_strokes(source.drawable_object_id, &[Some(rounded), Some(medium)])
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(source.drawable_object_id, 6)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_strokes(duplicate.drawable_object_id)
            .unwrap(),
        vec![Some(rounded), Some(medium)]
    );
    editor
        .set_body_chart_series_stroke(source.drawable_object_id, first, None)
        .unwrap();
    assert_eq!(
        editor
            .reset_body_chart_series_stroke(source.drawable_object_id, first)
            .unwrap(),
        None
    );

    let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_stroke(source.drawable_object_id, first)
            .unwrap(),
        None
    );
    assert_eq!(
        reopened
            .body_chart_series_stroke(source.drawable_object_id, second)
            .unwrap(),
        Some(medium)
    );
    assert_eq!(
        reopened
            .body_chart_series_strokes(duplicate.drawable_object_id)
            .unwrap(),
        vec![Some(rounded), Some(medium)]
    );
}

#[test]
fn scratch_document_supports_native_chart_shadow_crud() {
    let mut editor = PagesEditor::create_with_text("Chart shadow").unwrap();
    let source = editor
        .add_body_chart(
            "Chart shadow".encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let native_default = ChartShadow::native_default();
    assert_eq!(
        editor.body_chart_shadow(source.drawable_object_id).unwrap(),
        native_default
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_shadow(source.drawable_object_id, native_default)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let customized = chart_shadow();
    editor
        .set_body_chart_shadow(source.drawable_object_id, customized)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_shadow(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    editor
        .set_body_chart_shadow(source.drawable_object_id, ChartShadow::None)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_shadow(source.drawable_object_id)
            .unwrap(),
        ChartShadow::None
    );
    assert_eq!(
        reopened
            .body_chart_shadow(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    reopened
        .set_body_chart_shadow(duplicate.drawable_object_id, native_default)
        .unwrap();
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_pie_start_angle_crud() {
    let mut editor = PagesEditor::create_with_text("Pie rotation").unwrap();
    let source = editor
        .add_body_chart(
            "Pie rotation".encode_utf16().count(),
            ChartKind::Pie2d,
            pie_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_start_angle(source.drawable_object_id)
            .unwrap(),
        ChartPieStartAngle::ZERO
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_pie_start_angle(source.drawable_object_id, ChartPieStartAngle::ZERO)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let customized = ChartPieStartAngle::from_degrees(123.0).unwrap();
    editor
        .set_body_chart_pie_start_angle(source.drawable_object_id, customized)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    editor
        .set_body_chart_kind(duplicate.drawable_object_id, ChartKind::Donut2d)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_start_angle(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    editor
        .set_body_chart_pie_start_angle(source.drawable_object_id, ChartPieStartAngle::HALF_TURN)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_pie_start_angle(source.drawable_object_id)
            .unwrap(),
        ChartPieStartAngle::HALF_TURN
    );
    assert_eq!(
        reopened
            .body_chart_pie_start_angle(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    reopened
        .set_body_chart_pie_start_angle(duplicate.drawable_object_id, ChartPieStartAngle::ZERO)
        .unwrap();

    let column = reopened
        .add_body_chart(
            reopened.body_text().unwrap().encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let before_rejected_update = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .body_chart_pie_start_angle(column.drawable_object_id)
            .is_err()
    );
    assert!(
        reopened
            .set_body_chart_pie_start_angle(
                column.drawable_object_id,
                ChartPieStartAngle::QUARTER_TURN,
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected_update);

    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(column.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_donut_inner_radius_crud() {
    let mut editor = PagesEditor::create_with_text("Donut radius").unwrap();
    let source = editor
        .add_body_chart(
            "Donut radius".encode_utf16().count(),
            ChartKind::Donut2d,
            pie_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_donut_inner_radius(source.drawable_object_id)
            .unwrap(),
        ChartDonutInnerRadius::DEFAULT
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_donut_inner_radius(
            source.drawable_object_id,
            ChartDonutInnerRadius::DEFAULT,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let customized = ChartDonutInnerRadius::from_percent(42.0).unwrap();
    editor
        .set_body_chart_donut_inner_radius(source.drawable_object_id, customized)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_donut_inner_radius(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    editor
        .set_body_chart_kind(duplicate.drawable_object_id, ChartKind::Pie2d)
        .unwrap();
    let before_rejected_update = editor.to_bytes().unwrap();
    assert!(
        editor
            .body_chart_donut_inner_radius(duplicate.drawable_object_id)
            .is_err()
    );
    assert!(
        editor
            .set_body_chart_donut_inner_radius(
                duplicate.drawable_object_id,
                ChartDonutInnerRadius::MAXIMUM,
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_rejected_update);
    editor
        .set_body_chart_kind(duplicate.drawable_object_id, ChartKind::Donut3d)
        .unwrap();

    editor
        .set_body_chart_donut_inner_radius(
            source.drawable_object_id,
            ChartDonutInnerRadius::MINIMUM,
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_donut_inner_radius(source.drawable_object_id)
            .unwrap(),
        ChartDonutInnerRadius::MINIMUM
    );
    assert_eq!(
        reopened
            .body_chart_donut_inner_radius(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    reopened
        .set_body_chart_donut_inner_radius(
            duplicate.drawable_object_id,
            ChartDonutInnerRadius::DEFAULT,
        )
        .unwrap();
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_pie_wedge_explosion_crud() {
    let mut editor = PagesEditor::create_with_text("Revenue").unwrap();
    let source = editor
        .add_body_chart(7, ChartKind::Pie2d, pie_data(), POSITION, SIZE)
        .unwrap();
    let zeros = vec![ChartPieWedgeExplosion::ZERO; 3];
    assert_eq!(
        editor
            .body_chart_pie_wedge_explosions(source.drawable_object_id)
            .unwrap(),
        zeros
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_pie_wedge_explosions(source.drawable_object_id, &zeros)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let customized = [
        ChartPieWedgeExplosion::from_percent(10.0).unwrap(),
        ChartPieWedgeExplosion::from_percent(25.0).unwrap(),
        ChartPieWedgeExplosion::from_percent(40.0).unwrap(),
    ];
    editor
        .set_body_chart_pie_wedge_explosions(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_wedge_explosion(
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(1),
            )
            .unwrap(),
        customized[1]
    );
    editor
        .set_body_chart_pie_wedge_explosions(source.drawable_object_id, &zeros)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_pie_wedge_explosions(source.drawable_object_id, &customized)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(source.drawable_object_id, 7)
        .unwrap();
    editor
        .set_body_chart_kind(duplicate.drawable_object_id, ChartKind::Donut2d)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_wedge_explosions(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    let isolated = ChartPieWedgeExplosion::from_percent(55.0).unwrap();
    editor
        .set_body_chart_pie_wedge_explosion(
            source.drawable_object_id,
            ChartPieWedgeIndex::from_zero_based(0),
            isolated,
        )
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_pie_wedge_explosion(
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(0),
            )
            .unwrap(),
        isolated
    );
    assert_eq!(
        reopened
            .body_chart_pie_wedge_explosions(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected_updates = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_pie_wedge_explosions(source.drawable_object_id, &customized[..2],)
            .is_err()
    );
    assert!(
        reopened
            .set_body_chart_pie_wedge_explosion(
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(3),
                isolated,
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected_updates);

    let column = reopened
        .add_body_chart(7, ChartKind::Column2d, sample_data(), POSITION, SIZE)
        .unwrap();
    let before_wrong_kind = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .body_chart_pie_wedge_explosions(column.drawable_object_id)
            .is_err()
    );
    assert!(
        reopened
            .set_body_chart_pie_wedge_explosions(column.drawable_object_id, &customized,)
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_wrong_kind);

    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(column.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_pie_label_visibility_crud() {
    let mut editor = PagesEditor::create_with_text("Revenue").unwrap();
    let source = editor
        .add_body_chart(7, ChartKind::Pie2d, pie_data(), POSITION, SIZE)
        .unwrap();
    let defaults = vec![ChartPieLabelVisibility::DEFAULT; 3];
    let customized = [
        ChartPieLabelVisibility::DATA_POINT_NAMES_ONLY,
        ChartPieLabelVisibility::ALL,
        ChartPieLabelVisibility::HIDDEN,
    ];
    assert_eq!(
        editor
            .body_chart_pie_label_visibilities(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_pie_label_visibilities(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_pie_label_visibilities(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_label_visibility(
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
        .set_body_chart_pie_wedge_explosions(source.drawable_object_id, &explosions)
        .unwrap();
    editor
        .set_body_chart_pie_label_visibilities(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_wedge_explosions(source.drawable_object_id)
            .unwrap(),
        explosions
    );
    editor
        .set_body_chart_pie_wedge_explosions(
            source.drawable_object_id,
            &[ChartPieWedgeExplosion::ZERO; 3],
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_pie_label_visibilities(source.drawable_object_id, &customized)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(source.drawable_object_id, 7)
        .unwrap();
    editor
        .set_body_chart_kind(duplicate.drawable_object_id, ChartKind::Donut2d)
        .unwrap();
    editor
        .set_body_chart_pie_label_visibility(
            source.drawable_object_id,
            ChartPieWedgeIndex::from_zero_based(0),
            ChartPieLabelVisibility::VALUES_ONLY,
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_pie_label_visibilities(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_pie_label_visibilities(source.drawable_object_id, &customized[..2],)
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_pie_label_distance_crud() {
    let mut editor = PagesEditor::create_with_text("Revenue").unwrap();
    let source = editor
        .add_body_chart(7, ChartKind::Pie2d, pie_data(), POSITION, SIZE)
        .unwrap();
    let defaults = vec![ChartPieLabelDistance::DEFAULT; 3];
    let customized = [
        ChartPieLabelDistance::MINIMUM,
        ChartPieLabelDistance::from_percent(100.0).unwrap(),
        ChartPieLabelDistance::MAXIMUM,
    ];
    assert_eq!(
        editor
            .body_chart_pie_label_distances(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_pie_label_distances(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    let leader_line_defaults = [ChartPieLeaderLineVisibility::Visible; 3];
    let leader_line_customized = [
        ChartPieLeaderLineVisibility::Hidden,
        ChartPieLeaderLineVisibility::Visible,
        ChartPieLeaderLineVisibility::Hidden,
    ];
    assert_eq!(
        editor
            .body_chart_pie_leader_line_visibilities(source.drawable_object_id)
            .unwrap(),
        leader_line_defaults
    );
    editor
        .set_body_chart_pie_leader_line_visibilities(
            source.drawable_object_id,
            &leader_line_defaults,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_body_chart_pie_leader_line_visibilities(
            source.drawable_object_id,
            &leader_line_customized,
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_leader_line_visibility(
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(0),
            )
            .unwrap(),
        ChartPieLeaderLineVisibility::Hidden
    );
    editor
        .set_body_chart_pie_leader_line_visibilities(
            source.drawable_object_id,
            &leader_line_defaults,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_pie_label_distances(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_label_distance(
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
        .set_body_chart_pie_label_visibilities(source.drawable_object_id, &visibilities)
        .unwrap();
    editor
        .set_body_chart_pie_label_distances(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_pie_label_visibilities(source.drawable_object_id)
            .unwrap(),
        visibilities
    );
    editor
        .set_body_chart_pie_label_visibilities(
            source.drawable_object_id,
            &[ChartPieLabelVisibility::DEFAULT; 3],
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_pie_label_distances(source.drawable_object_id, &customized)
        .unwrap();
    editor
        .set_body_chart_pie_leader_line_visibilities(
            source.drawable_object_id,
            &leader_line_customized,
        )
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(source.drawable_object_id, 7)
        .unwrap();
    editor
        .set_body_chart_kind(duplicate.drawable_object_id, ChartKind::Donut2d)
        .unwrap();
    editor
        .set_body_chart_pie_label_distance(
            source.drawable_object_id,
            ChartPieWedgeIndex::from_zero_based(0),
            ChartPieLabelDistance::DEFAULT,
        )
        .unwrap();
    editor
        .set_body_chart_pie_leader_line_visibility(
            source.drawable_object_id,
            ChartPieWedgeIndex::from_zero_based(0),
            ChartPieLeaderLineVisibility::Visible,
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_pie_label_distances(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    assert_eq!(
        reopened
            .body_chart_pie_leader_line_visibilities(duplicate.drawable_object_id)
            .unwrap(),
        leader_line_customized
    );
    assert_eq!(
        reopened
            .body_chart_pie_leader_line_visibilities(source.drawable_object_id)
            .unwrap(),
        [
            ChartPieLeaderLineVisibility::Visible,
            ChartPieLeaderLineVisibility::Visible,
            ChartPieLeaderLineVisibility::Hidden,
        ]
    );
    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_pie_label_distances(source.drawable_object_id, &customized[..2],)
            .is_err()
    );
    assert!(
        reopened
            .set_body_chart_pie_leader_line_visibilities(
                source.drawable_object_id,
                &leader_line_customized[..2],
            )
            .is_err()
    );
    assert!(
        reopened
            .body_chart_pie_leader_line_visibility(
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(3),
            )
            .is_err()
    );
    assert!(
        reopened
            .body_chart_pie_label_distance(
                source.drawable_object_id,
                ChartPieWedgeIndex::from_zero_based(3),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_series_value_label_crud() {
    const BODY_TEXT: &str = "Chart value labels";

    let mut editor = PagesEditor::create_with_text(BODY_TEXT).unwrap();
    let source = editor
        .add_body_chart(
            BODY_TEXT.encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let defaults = [ChartSeriesValueLabelVisibility::Hidden; 2];
    let customized = [
        ChartSeriesValueLabelVisibility::Visible,
        ChartSeriesValueLabelVisibility::Hidden,
    ];

    assert_eq!(
        editor
            .body_chart_series_value_label_visibilities(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_value_label_visibilities(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_series_value_label_visibilities(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_value_label_visibility(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(0),
            )
            .unwrap(),
        ChartSeriesValueLabelVisibility::Visible
    );
    editor
        .set_body_chart_series_value_label_visibilities(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_series_value_label_visibilities(source.drawable_object_id, &customized)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    editor
        .set_body_chart_series_value_label_visibility(
            source.drawable_object_id,
            ChartSeriesIndex::from_zero_based(0),
            ChartSeriesValueLabelVisibility::Hidden,
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_value_label_visibilities(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .body_chart_series_value_label_visibilities(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_series_value_label_visibilities(
                source.drawable_object_id,
                &customized[..1],
            )
            .is_err()
    );
    assert!(
        reopened
            .body_chart_series_value_label_visibility(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_series_value_label_location_crud() {
    const BODY_TEXT: &str = "Chart value-label Location";

    let mut editor = PagesEditor::create_with_text(BODY_TEXT).unwrap();
    let source = editor
        .add_body_chart(
            BODY_TEXT.encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let defaults = [ChartSeriesValueLabelLocation::Top; 2];
    let customized = [
        ChartSeriesValueLabelLocation::Outside,
        ChartSeriesValueLabelLocation::Top,
    ];

    assert_eq!(
        editor
            .body_chart_series_value_label_locations(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_value_label_locations(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_series_value_label_locations(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_value_label_location(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(0),
            )
            .unwrap(),
        ChartSeriesValueLabelLocation::Outside
    );
    editor
        .set_body_chart_series_value_label_locations(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_series_value_label_locations(source.drawable_object_id, &customized)
        .unwrap();
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    editor
        .set_body_chart_series_value_label_location(
            source.drawable_object_id,
            ChartSeriesIndex::from_zero_based(0),
            ChartSeriesValueLabelLocation::Top,
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_value_label_locations(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .body_chart_series_value_label_locations(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_series_value_label_locations(
                source.drawable_object_id,
                &customized[..1],
            )
            .is_err()
    );
    assert!(
        reopened
            .body_chart_series_value_label_location(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_series_value_label_affix_crud() {
    const BODY_TEXT: &str = "Chart value-label affixes";

    let mut editor = PagesEditor::create_with_text(BODY_TEXT).unwrap();
    let source = editor
        .add_body_chart(
            BODY_TEXT.encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let defaults = vec![ChartSeriesValueLabelAffixes::default(); 2];
    let customized = vec![
        ChartSeriesValueLabelAffixes::new("$", " USD"),
        ChartSeriesValueLabelAffixes::new("€", " net"),
    ];

    assert_eq!(
        editor
            .body_chart_series_value_label_affixes(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_value_label_affixes(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_series_value_label_affixes(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_value_label_affix(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(0),
            )
            .unwrap()
            .suffix(),
        " USD"
    );
    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for series in 0..2 {
        editor
            .set_body_chart_series_value_label_affix(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(series),
                ChartSeriesValueLabelAffixes::default(),
            )
            .unwrap();
    }

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_value_label_affixes(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .body_chart_series_value_label_affixes(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_series_value_label_affixes(source.drawable_object_id, &customized[..1],)
            .is_err()
    );
    assert!(
        reopened
            .body_chart_series_value_label_affix(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_series_value_label_number_format_crud() {
    const BODY_TEXT: &str = "Chart value-label number formats";

    let mut editor = PagesEditor::create_with_text(BODY_TEXT).unwrap();
    let source = editor
        .add_body_chart(
            BODY_TEXT.encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
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
            .body_chart_series_value_label_number_formats(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_value_label_number_formats(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_body_chart_series_value_label_number_formats(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_value_label_number_format(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(0),
            )
            .unwrap(),
        fixed_two
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    editor
        .set_body_chart_series_value_label_number_format(
            source.drawable_object_id,
            ChartSeriesIndex::from_zero_based(0),
            ChartSeriesValueLabelNumberFormat::NATIVE_DEFAULT,
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_value_label_number_formats(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .body_chart_series_value_label_number_formats(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_series_value_label_number_formats(
                source.drawable_object_id,
                &customized[..1],
            )
            .is_err()
    );
    assert!(
        reopened
            .body_chart_series_value_label_number_format(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_series_value_label_auto_fit_crud() {
    const BODY_TEXT: &str = "Chart value-label Auto-Fit";

    let mut editor = PagesEditor::create_with_text(BODY_TEXT).unwrap();
    let source = editor
        .add_body_chart(
            BODY_TEXT.encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
        .unwrap();
    let defaults = vec![ChartSeriesValueLabelAutoFit::Enabled; 2];
    let customized = vec![
        ChartSeriesValueLabelAutoFit::Disabled,
        ChartSeriesValueLabelAutoFit::Enabled,
    ];

    assert_eq!(
        editor
            .body_chart_series_value_label_auto_fits(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_value_label_auto_fits(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_body_chart_series_value_label_auto_fits(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_value_label_auto_fit(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(0),
            )
            .unwrap(),
        ChartSeriesValueLabelAutoFit::Disabled
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    editor
        .set_body_chart_series_value_label_auto_fit(
            source.drawable_object_id,
            ChartSeriesIndex::from_zero_based(0),
            ChartSeriesValueLabelAutoFit::Enabled,
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_value_label_auto_fits(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .body_chart_series_value_label_auto_fits(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_series_value_label_auto_fits(
                source.drawable_object_id,
                &customized[..1],
            )
            .is_err()
    );
    assert!(
        reopened
            .body_chart_series_value_label_auto_fit(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_series_trendline_crud() {
    const BODY_TEXT: &str = "Chart series trendlines";

    let mut editor = PagesEditor::create_with_text(BODY_TEXT).unwrap();
    let source = editor
        .add_body_chart(
            BODY_TEXT.encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
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
            .body_chart_series_trendlines(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_trendlines(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_body_chart_series_trendlines(source.drawable_object_id, &customized)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_trendline(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(1),
            )
            .unwrap(),
        customized[1]
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for series in 0..2 {
        editor
            .set_body_chart_series_trendline(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(series),
                ChartSeriesTrendline::none(),
            )
            .unwrap();
    }
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_trendlines(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .body_chart_series_trendlines(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_series_trendlines(source.drawable_object_id, &customized[..1])
            .is_err()
    );
    assert!(
        reopened
            .body_chart_series_trendline(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(2),
            )
            .is_err()
    );
    assert!(ChartSeriesTrendline::unsupported(1).is_err());
    assert!(ChartSeriesTrendlinePolynomialOrder::new(7).is_err());
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}

#[test]
fn scratch_document_supports_native_series_error_bar_crud() {
    const BODY_TEXT: &str = "Chart series error bars";

    let mut editor = PagesEditor::create_with_text(BODY_TEXT).unwrap();
    let source = editor
        .add_body_chart(
            BODY_TEXT.encode_utf16().count(),
            ChartKind::Column2d,
            sample_data(),
            POSITION,
            SIZE,
        )
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
            .body_chart_series_error_bars(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        editor
            .body_chart_series_error_bar_auto_fits(source.drawable_object_id)
            .unwrap(),
        default_auto_fits
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_series_error_bars(source.drawable_object_id, &defaults)
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    editor
        .set_body_chart_series_error_bars(source.drawable_object_id, &customized)
        .unwrap();
    editor
        .set_body_chart_series_error_bar_auto_fits(source.drawable_object_id, &customized_auto_fits)
        .unwrap();
    assert_eq!(
        editor
            .body_chart_series_error_bar(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(1),
            )
            .unwrap(),
        customized[1]
    );
    assert_eq!(
        editor
            .body_chart_series_error_bar_auto_fit(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(0),
            )
            .unwrap(),
        ChartSeriesErrorBarAutoFit::Disabled
    );

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for series in 0..2 {
        editor
            .set_body_chart_series_error_bar(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(series),
                ChartSeriesErrorBars::None,
            )
            .unwrap();
    }
    editor
        .set_body_chart_series_error_bar_auto_fits(source.drawable_object_id, &default_auto_fits)
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .body_chart_series_error_bars(source.drawable_object_id)
            .unwrap(),
        defaults
    );
    assert_eq!(
        reopened
            .body_chart_series_error_bars(duplicate.drawable_object_id)
            .unwrap(),
        customized
    );
    assert_eq!(
        reopened
            .body_chart_series_error_bar_auto_fits(source.drawable_object_id)
            .unwrap(),
        default_auto_fits
    );
    assert_eq!(
        reopened
            .body_chart_series_error_bar_auto_fits(duplicate.drawable_object_id)
            .unwrap(),
        customized_auto_fits
    );

    let before_rejected = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_body_chart_series_error_bars(source.drawable_object_id, &customized[..1])
            .is_err()
    );
    assert!(
        reopened
            .set_body_chart_series_error_bar_auto_fits(
                source.drawable_object_id,
                &customized_auto_fits[..1],
            )
            .is_err()
    );
    assert!(
        reopened
            .body_chart_series_error_bar(
                source.drawable_object_id,
                ChartSeriesIndex::from_zero_based(2),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_rejected);
    reopened
        .remove_body_chart(source.drawable_object_id)
        .unwrap();
    reopened
        .remove_body_chart(duplicate.drawable_object_id)
        .unwrap();
    assert!(reopened.body_charts().unwrap().is_empty());
}
