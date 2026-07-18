//! Source-built Pages chart CRUD regression tests.

use super::*;

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
