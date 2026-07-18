use super::*;
use crate::keynote::KeynoteDocumentBuilder;

fn table_geometry() -> (DrawablePoint, DrawableSize) {
    (
        DrawablePoint { x: 120.0, y: 180.0 },
        DrawableSize {
            width: 840.0,
            height: 360.0,
        },
    )
}

#[test]
fn source_built_table_roundtrips_full_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    assert!(editor.slide_tables(0).unwrap().is_empty());
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Forecast", 3, 3, position, size)
        .unwrap();
    assert_eq!(table.name, "Forecast");
    assert_eq!((table.rows, table.columns), (3, 3));

    editor
        .set_slide_table_cell(
            0,
            table.model_object_id,
            0,
            0,
            KeynoteTableCellValue::Text("Region".to_owned()),
        )
        .unwrap();
    editor
        .set_slide_table_cell(
            0,
            table.model_object_id,
            1,
            1,
            KeynoteTableCellValue::Number(42.5),
        )
        .unwrap();
    editor
        .rename_slide_table(0, table.model_object_id, "Outlook")
        .unwrap();
    editor
        .resize_slide_table(0, table.model_object_id, 4, 4)
        .unwrap();
    let replacement = DrawableGeometry {
        position: Some(DrawablePoint { x: 180.0, y: 210.0 }),
        size: Some(DrawableSize {
            width: 900.0,
            height: 420.0,
        }),
        flags: Some(TABLE_GEOMETRY_FLAGS),
        angle: Some(TABLE_ANGLE_DEGREES),
    };
    editor
        .set_slide_table_geometry(0, table.drawable_object_id, replacement)
        .unwrap();
    let headers = KeynoteTableHeaderSettings {
        header_rows: Some(KeynoteTableHeaderCount::TWO),
        header_columns: None,
        footer_rows: Some(KeynoteTableHeaderCount::ONE),
        ..Default::default()
    };
    editor
        .set_slide_table_header_settings(0, table.model_object_id, headers)
        .unwrap();
    editor
        .set_slide_table_column_width(
            0,
            table.model_object_id,
            0,
            KeynoteTableDimensionSize::points(300.0).unwrap(),
        )
        .unwrap();
    editor
        .set_slide_table_row_height(
            0,
            table.model_object_id,
            0,
            KeynoteTableDimensionSize::points(140.0).unwrap(),
        )
        .unwrap();
    let laid_out_geometry = DrawableGeometry {
        size: Some(DrawableSize {
            width: 975.0,
            height: 455.0,
        }),
        ..replacement
    };

    let bytes = editor.to_bytes().unwrap();
    let mut reopened = KeynoteEditor::from_bytes(&bytes).unwrap();
    let materialized = reopened.slide_table(0, table.model_object_id).unwrap();
    assert_eq!(materialized.info.name, "Outlook");
    assert_eq!((materialized.info.rows, materialized.info.columns), (4, 4));
    assert_eq!(materialized.info.geometry, laid_out_geometry);
    assert_eq!(
        reopened
            .slide_table_header_settings(0, table.model_object_id)
            .unwrap(),
        headers
    );
    assert_eq!(
        reopened
            .slide_table_column_width(0, table.model_object_id, 0)
            .unwrap(),
        KeynoteTableDimensionSize::points(300.0).unwrap()
    );
    assert_eq!(
        reopened
            .slide_table_row_height(0, table.model_object_id, 0)
            .unwrap(),
        KeynoteTableDimensionSize::points(140.0).unwrap()
    );
    assert_eq!(
        materialized.get_cell(0, 0),
        Some(&KeynoteTableCellValue::Text("Region".to_owned()))
    );
    assert_eq!(
        materialized.get_cell(1, 1),
        Some(&KeynoteTableCellValue::Number(42.5))
    );

    reopened
        .clear_slide_table_cell(0, table.model_object_id, 1, 1)
        .unwrap();
    assert!(
        reopened
            .slide_table(0, table.model_object_id)
            .unwrap()
            .get_cell(1, 1)
            .is_none_or(KeynoteTableCellValue::is_empty)
    );
    let removed = reopened
        .remove_slide_table(0, table.drawable_object_id)
        .unwrap();
    assert_eq!(removed.table.model_object_id, table.model_object_id);
    assert!(reopened.slide_tables(0).unwrap().is_empty());
    assert!(KeynoteEditor::from_bytes(&reopened.to_bytes().unwrap()).is_ok());
}

#[test]
fn source_built_table_roundtrips_formula_crud_transactionally() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Formula", 3, 2, position, size)
        .unwrap();
    editor
        .set_slide_table_formula(
            0,
            table.model_object_id,
            2,
            1,
            KeynoteTableFormulaExpression::function(
                "SUM",
                [
                    KeynoteTableFormulaExpression::Number(1.0),
                    KeynoteTableFormulaExpression::Number(2.0),
                ],
            ),
            KeynoteTableFormulaCachedValue::Number(3.0),
        )
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_formula(0, table.model_object_id, 2, 1)
            .unwrap()
            .as_deref(),
        Some("=SUM(1,2)")
    );
    reopened
        .set_slide_table_formula(
            0,
            table.model_object_id,
            2,
            1,
            KeynoteTableFormulaExpression::function(
                "SUM",
                [
                    KeynoteTableFormulaExpression::Number(3.0),
                    KeynoteTableFormulaExpression::Number(4.0),
                ],
            ),
            KeynoteTableFormulaCachedValue::Number(7.0),
        )
        .unwrap();
    assert_eq!(
        reopened
            .slide_table_formula(0, table.model_object_id, 2, 1)
            .unwrap()
            .as_deref(),
        Some("=SUM(3,4)")
    );

    let before = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_slide_table_formula(
                0,
                table.model_object_id,
                usize::MAX,
                1,
                KeynoteTableFormulaExpression::Number(1.0),
                KeynoteTableFormulaCachedValue::Number(1.0),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before);
    assert_eq!(
        reopened
            .clear_slide_table_formula(0, table.model_object_id, 2, 1)
            .unwrap(),
        "=SUM(3,4)"
    );
    assert_eq!(
        reopened
            .slide_table_formula(0, table.model_object_id, 2, 1)
            .unwrap(),
        None
    );
    let cleared = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .clear_slide_table_formula(0, table.model_object_id, 2, 1)
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), cleared);
}

#[test]
fn tables_on_multiple_slides_remain_isolated() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let layout = editor
        .slide_layouts()
        .unwrap()
        .into_iter()
        .find(|layout| layout.is_default)
        .unwrap();
    editor.add_slide(layout.id).unwrap();
    let (position, size) = table_geometry();
    let first = editor
        .add_slide_table(0, "First", 2, 2, position, size)
        .unwrap();
    let second = editor
        .add_slide_table(1, "Second", 2, 2, position, size)
        .unwrap();
    assert_ne!(first.drawable_object_id, second.drawable_object_id);
    assert_ne!(first.model_object_id, second.model_object_id);
    assert_eq!(editor.slide_tables(0).unwrap()[0].name, "First");
    assert_eq!(editor.slide_tables(1).unwrap()[0].name, "Second");
}
