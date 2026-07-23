//! Source-built Keynote chart CRUD regression tests.

use super::*;
use crate::charts::ChartAxis;
use crate::keynote::KeynoteDocumentBuilder;

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
