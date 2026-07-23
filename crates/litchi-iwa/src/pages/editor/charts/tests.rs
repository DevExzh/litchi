//! Source-built Pages chart CRUD regression tests.

use super::*;
use crate::charts::{
    ChartAxis, ChartAxisBound, ChartAxisMajorStepCount, ChartAxisMinorStepCount,
    ChartValueAxisBounds, ChartValueAxisSteps,
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

    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert_eq!(
            editor
                .body_chart_axis_title(source.drawable_object_id, axis)
                .unwrap(),
            None
        );
    }
    editor
        .set_body_chart_axis_title(source.drawable_object_id, ChartAxis::Category, "Month")
        .unwrap();
    editor
        .set_body_chart_axis_title(source.drawable_object_id, ChartAxis::Value, "Revenue")
        .unwrap();

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for (axis, title) in [
        (ChartAxis::Category, "Month"),
        (ChartAxis::Value, "Revenue"),
    ] {
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
        .set_body_chart_axis_title(
            source.drawable_object_id,
            ChartAxis::Category,
            "Updated month",
        )
        .unwrap();
    editor
        .set_body_chart_axis_title(
            source.drawable_object_id,
            ChartAxis::Value,
            "Updated revenue",
        )
        .unwrap();
    assert_eq!(
        editor
            .body_chart_axis_title(source.drawable_object_id, ChartAxis::Category)
            .unwrap()
            .as_deref(),
        Some("Updated month")
    );
    assert_eq!(
        editor
            .body_chart_axis_title(source.drawable_object_id, ChartAxis::Value)
            .unwrap()
            .as_deref(),
        Some("Updated revenue")
    );
    assert_eq!(
        editor
            .body_chart_axis_title(duplicate.drawable_object_id, ChartAxis::Category)
            .unwrap()
            .as_deref(),
        Some("Month")
    );
    assert_eq!(
        editor
            .body_chart_axis_title(duplicate.drawable_object_id, ChartAxis::Value)
            .unwrap()
            .as_deref(),
        Some("Revenue")
    );

    for axis in [ChartAxis::Category, ChartAxis::Value] {
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
            .body_chart_axis_title(duplicate.drawable_object_id, ChartAxis::Category)
            .unwrap()
            .as_deref(),
        Some("Month")
    );
    assert_eq!(
        reopened
            .body_chart_axis_title(duplicate.drawable_object_id, ChartAxis::Value)
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

    for axis in [ChartAxis::Category, ChartAxis::Value] {
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
    for axis in [ChartAxis::Category, ChartAxis::Value] {
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
    for axis in [ChartAxis::Category, ChartAxis::Value] {
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
            .body_chart_axis_major_gridlines_visible(source.drawable_object_id, ChartAxis::Category)
            .unwrap()
    );
    assert!(
        editor
            .body_chart_axis_major_gridlines_visible(source.drawable_object_id, ChartAxis::Value)
            .unwrap()
    );
    let baseline = editor.to_bytes().unwrap();
    editor
        .set_body_chart_axis_major_gridlines_visible(
            source.drawable_object_id,
            ChartAxis::Category,
            false,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_body_chart_axis_major_gridlines_visible(
            source.drawable_object_id,
            ChartAxis::Category,
            true,
        )
        .unwrap();
    editor
        .set_body_chart_axis_major_gridlines_visible(
            source.drawable_object_id,
            ChartAxis::Value,
            false,
        )
        .unwrap();

    let duplicate = editor
        .duplicate_body_chart(
            source.drawable_object_id,
            editor.body_text().unwrap().encode_utf16().count(),
        )
        .unwrap();
    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert_eq!(
            editor
                .body_chart_axis_major_gridlines_visible(duplicate.drawable_object_id, axis)
                .unwrap(),
            axis == ChartAxis::Category
        );
    }

    editor
        .set_body_chart_axis_major_gridlines_visible(
            source.drawable_object_id,
            ChartAxis::Category,
            false,
        )
        .unwrap();
    editor
        .set_body_chart_axis_major_gridlines_visible(
            source.drawable_object_id,
            ChartAxis::Value,
            true,
        )
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    for axis in [ChartAxis::Category, ChartAxis::Value] {
        assert_eq!(
            reopened
                .body_chart_axis_major_gridlines_visible(source.drawable_object_id, axis)
                .unwrap(),
            axis == ChartAxis::Value
        );
        assert_eq!(
            reopened
                .body_chart_axis_major_gridlines_visible(duplicate.drawable_object_id, axis)
                .unwrap(),
            axis == ChartAxis::Category
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

    for axis in [ChartAxis::Category, ChartAxis::Value] {
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
            ChartAxis::Category,
            false,
        )
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    for axis in [ChartAxis::Category, ChartAxis::Value] {
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
    for axis in [ChartAxis::Category, ChartAxis::Value] {
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
    for axis in [ChartAxis::Category, ChartAxis::Value] {
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
