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
    let editor = KeynoteDocumentBuilder::new().build().unwrap();
    assert!(editor.slide_tables(0).unwrap().is_empty());
    let mut package = editor.into_package();
    let engine = package.remove_entry("Index/CalculationEngine.iwa").unwrap();
    package
        .insert_entry("Index/CalculationEngine-81.iwa", engine)
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Forecast", 3, 3, position, size)
        .unwrap();
    assert_eq!(table.name, "Forecast");
    assert_eq!((table.rows, table.columns), (3, 3));

    assert_eq!(
        editor
            .set_slide_table_cells(
                0,
                table.model_object_id,
                [
                    KeynoteTableCellUpdate::new(
                        0,
                        0,
                        KeynoteTableCellValue::Text("Region".to_owned()),
                    ),
                    KeynoteTableCellUpdate::new(1, 1, KeynoteTableCellValue::Number(42.5)),
                ],
            )
            .unwrap(),
        2
    );
    let before_invalid_batch = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_slide_table_cells(
                0,
                table.model_object_id,
                [
                    KeynoteTableCellUpdate::new(2, 0, KeynoteTableCellValue::Boolean(true),),
                    KeynoteTableCellUpdate::clear(2, 0),
                ],
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_invalid_batch);
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
fn source_built_table_duplication_clones_formula_storage_and_geometry() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let source = editor
        .add_slide_table(0, "Budget", 3, 2, position, size)
        .unwrap();
    editor
        .set_slide_table_cells(
            0,
            source.model_object_id,
            [
                KeynoteTableCellUpdate::new(
                    0,
                    0,
                    KeynoteTableCellValue::Text("Category".to_owned()),
                ),
                KeynoteTableCellUpdate::new(1, 0, KeynoteTableCellValue::Text("Travel".to_owned())),
                KeynoteTableCellUpdate::new(1, 1, KeynoteTableCellValue::Number(125.0)),
            ],
        )
        .unwrap();
    editor
        .set_slide_table_formula(
            0,
            source.model_object_id,
            2,
            1,
            KeynoteTableFormulaExpression::function(
                "SUM",
                [
                    KeynoteTableFormulaExpression::Number(100.0),
                    KeynoteTableFormulaExpression::Number(25.0),
                ],
            ),
            KeynoteTableFormulaCachedValue::Number(125.0),
        )
        .unwrap();

    let copied = editor
        .duplicate_slide_table(0, source.drawable_object_id)
        .unwrap();
    assert_ne!(copied.drawable_object_id, source.drawable_object_id);
    assert_ne!(copied.model_object_id, source.model_object_id);
    assert_eq!(copied.name, "Budget copy");
    assert_eq!((copied.rows, copied.columns), (source.rows, source.columns));
    let mut expected_geometry = source.geometry;
    if let Some(position) = expected_geometry.position.as_mut() {
        position.x += TABLE_DUPLICATE_OFFSET;
        position.y += TABLE_DUPLICATE_OFFSET;
    }
    assert_eq!(copied.geometry, expected_geometry);
    assert_eq!(
        editor
            .slide_table(copied.slide_index, copied.model_object_id)
            .unwrap()
            .get_cell(1, 0),
        Some(&KeynoteTableCellValue::Text("Travel".to_owned()))
    );
    assert_eq!(
        editor
            .slide_table_formula(copied.slide_index, copied.model_object_id, 2, 1)
            .unwrap()
            .as_deref(),
        Some("=SUM(100,25)")
    );

    editor
        .set_slide_table_cell(
            copied.slide_index,
            copied.model_object_id,
            1,
            0,
            KeynoteTableCellValue::Text("Lodging".to_owned()),
        )
        .unwrap();
    assert_eq!(
        editor
            .slide_table(source.slide_index, source.model_object_id)
            .unwrap()
            .get_cell(1, 0),
        Some(&KeynoteTableCellValue::Text("Travel".to_owned()))
    );
    assert_eq!(
        editor
            .duplicate_slide_table(0, source.drawable_object_id)
            .unwrap()
            .name,
        "Budget copy 2"
    );

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.slide_tables(0).unwrap().len(), 3);
    reopened
        .remove_slide_table(copied.slide_index, copied.drawable_object_id)
        .unwrap();
    assert_eq!(reopened.slide_tables(0).unwrap().len(), 2);
    assert_eq!(
        reopened
            .slide_table(source.slide_index, source.model_object_id)
            .unwrap()
            .get_cell(1, 0),
        Some(&KeynoteTableCellValue::Text("Travel".to_owned()))
    );
}

#[test]
fn source_built_table_roundtrips_full_table_sort_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Cities", 5, 2, position, size)
        .unwrap();
    let model_id = table.model_object_id;
    editor
        .set_slide_table_header_settings(
            0,
            model_id,
            KeynoteTableHeaderSettings {
                header_rows: Some(KeynoteTableHeaderCount::ONE),
                ..Default::default()
            },
        )
        .unwrap();
    editor
        .set_slide_table_cells(
            0,
            model_id,
            [
                KeynoteTableCellUpdate::new(0, 0, KeynoteTableCellValue::Text("Name".to_owned())),
                KeynoteTableCellUpdate::new(0, 1, KeynoteTableCellValue::Text("Marker".to_owned())),
                KeynoteTableCellUpdate::new(1, 0, KeynoteTableCellValue::Text("zebra".to_owned())),
                KeynoteTableCellUpdate::new(1, 1, KeynoteTableCellValue::Text("last".to_owned())),
                KeynoteTableCellUpdate::new(2, 0, KeynoteTableCellValue::Text("apple".to_owned())),
                KeynoteTableCellUpdate::new(
                    2,
                    1,
                    KeynoteTableCellValue::Text("first apple".to_owned()),
                ),
                KeynoteTableCellUpdate::new(3, 0, KeynoteTableCellValue::Text("banana".to_owned())),
                KeynoteTableCellUpdate::new(3, 1, KeynoteTableCellValue::Text("middle".to_owned())),
                KeynoteTableCellUpdate::new(4, 0, KeynoteTableCellValue::Text("apple".to_owned())),
                KeynoteTableCellUpdate::new(
                    4,
                    1,
                    KeynoteTableCellValue::Text("second apple".to_owned()),
                ),
            ],
        )
        .unwrap();
    let order = KeynoteTableSortOrder::new([KeynoteTableSortRule::new(
        KeynoteTableSortColumnIndex::new(0).unwrap(),
        KeynoteTableSortDirection::Ascending,
    )])
    .unwrap();

    assert_eq!(editor.slide_table_sort_order(0, model_id).unwrap(), None);
    editor
        .set_slide_table_sort_order(0, model_id, order.clone())
        .unwrap();
    assert_eq!(
        editor.slide_table_sort_order(0, model_id).unwrap(),
        Some(order.clone())
    );
    assert!(editor.apply_slide_table_sort_order(0, model_id).unwrap());
    assert!(!editor.apply_slide_table_sort_order(0, model_id).unwrap());
    let materialized = editor.slide_table(0, model_id).unwrap();
    assert_eq!(
        materialized.get_cell(0, 0),
        Some(&KeynoteTableCellValue::Text("Name".to_owned()))
    );
    assert_eq!(
        materialized.get_cell(1, 0),
        Some(&KeynoteTableCellValue::Text("apple".to_owned()))
    );
    assert_eq!(
        materialized.get_cell(1, 1),
        Some(&KeynoteTableCellValue::Text("first apple".to_owned()))
    );
    assert_eq!(
        materialized.get_cell(2, 1),
        Some(&KeynoteTableCellValue::Text("second apple".to_owned()))
    );
    assert_eq!(
        materialized.get_cell(3, 0),
        Some(&KeynoteTableCellValue::Text("banana".to_owned()))
    );
    assert_eq!(
        materialized.get_cell(4, 0),
        Some(&KeynoteTableCellValue::Text("zebra".to_owned()))
    );

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.slide_table_sort_order(0, model_id).unwrap(),
        Some(order.clone())
    );
    let unchanged = reopened.to_bytes().unwrap();
    reopened
        .set_slide_table_sort_order(0, model_id, order)
        .unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), unchanged);

    let invalid = KeynoteTableSortOrder::new([KeynoteTableSortRule::new(
        KeynoteTableSortColumnIndex::new(2).unwrap(),
        KeynoteTableSortDirection::Ascending,
    )])
    .unwrap();
    let before_invalid = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_slide_table_sort_order(0, model_id, invalid)
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_invalid);

    let selected_order = KeynoteTableSortOrder::selected_rows([KeynoteTableSortRule::new(
        KeynoteTableSortColumnIndex::new(0).unwrap(),
        KeynoteTableSortDirection::Descending,
    )])
    .unwrap();
    reopened
        .set_slide_table_sort_order(0, model_id, selected_order.clone())
        .unwrap();
    assert_eq!(
        reopened.slide_table_sort_order(0, model_id).unwrap(),
        Some(selected_order)
    );
    let before_wrong_executor = reopened.to_bytes().unwrap();
    assert!(reopened.apply_slide_table_sort_order(0, model_id).is_err());
    assert_eq!(reopened.to_bytes().unwrap(), before_wrong_executor);
    assert!(
        reopened
            .apply_slide_table_sort_order_to_rows(
                0,
                model_id,
                KeynoteTableSortRowRange::new(1, 4).unwrap(),
            )
            .unwrap()
    );
    let table = reopened.slide_table(0, model_id).unwrap();
    assert_eq!(
        table.get_cell(1, 1),
        Some(&KeynoteTableCellValue::Text("first apple".to_owned()))
    );
    assert_eq!(
        table.get_cell(2, 0),
        Some(&KeynoteTableCellValue::Text("zebra".to_owned()))
    );
    assert_eq!(
        table.get_cell(3, 0),
        Some(&KeynoteTableCellValue::Text("banana".to_owned()))
    );
    assert_eq!(
        table.get_cell(4, 1),
        Some(&KeynoteTableCellValue::Text("second apple".to_owned()))
    );

    reopened.clear_slide_table_sort_order(0, model_id).unwrap();
    assert_eq!(reopened.slide_table_sort_order(0, model_id).unwrap(), None);
    let unchanged = reopened.to_bytes().unwrap();
    reopened.clear_slide_table_sort_order(0, model_id).unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), unchanged);
    assert!(reopened.apply_slide_table_sort_order(0, model_id).is_err());
    assert_eq!(reopened.to_bytes().unwrap(), unchanged);
}

#[test]
fn source_built_table_roundtrips_formula_crud_transactionally() {
    let editor = KeynoteDocumentBuilder::new().build().unwrap();
    let mut package = editor.into_package();
    let engine = package.remove_entry("Index/CalculationEngine.iwa").unwrap();
    package
        .insert_entry("Index/CalculationEngine-82.iwa", engine)
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
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
fn source_built_table_roundtrips_section_relative_axis_crud_transactionally() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Topology", 4, 4, position, size)
        .unwrap();
    let model_id = table.model_object_id;
    let row_size = KeynoteTableDimensionSize::points(88.0).unwrap();
    let column_size = KeynoteTableDimensionSize::points(144.0).unwrap();
    editor
        .set_slide_table_cell(
            0,
            model_id,
            1,
            1,
            KeynoteTableCellValue::Text("shift me".to_owned()),
        )
        .unwrap();
    editor
        .set_slide_table_formula(
            0,
            model_id,
            2,
            2,
            KeynoteTableFormulaExpression::cell(KeynoteTableFormulaCellReference::relative(1, 1)),
            KeynoteTableFormulaCachedValue::Number(7.0),
        )
        .unwrap();
    editor
        .set_slide_table_row_height(0, model_id, 1, row_size)
        .unwrap();
    editor
        .set_slide_table_column_width(0, model_id, 1, column_size)
        .unwrap();
    let baseline_geometry = editor.slide_tables(0).unwrap()[0].geometry;
    let baseline = editor.to_bytes().unwrap();

    editor
        .insert_slide_table_row(0, model_id, KeynoteTableRowInsertion::body(1))
        .unwrap();
    editor
        .insert_slide_table_column(0, model_id, KeynoteTableColumnInsertion::body(1))
        .unwrap();
    let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let shifted = reopened.slide_table(0, model_id).unwrap();
    assert_eq!((shifted.info.rows, shifted.info.columns), (5, 5));
    assert_ne!(shifted.info.geometry.size, baseline_geometry.size);
    assert_eq!(
        shifted.get_cell(1, 1),
        Some(&KeynoteTableCellValue::Text("shift me".to_owned()))
    );
    assert_eq!(
        shifted.get_cell(3, 3),
        Some(&KeynoteTableCellValue::Formula("=B2".to_owned()))
    );
    assert_eq!(
        reopened.slide_table_row_height(0, model_id, 1).unwrap(),
        row_size
    );
    assert_eq!(
        reopened.slide_table_column_width(0, model_id, 1).unwrap(),
        column_size
    );

    editor
        .remove_slide_table_column(0, model_id, KeynoteTableColumnDeletion::body(1))
        .unwrap();
    editor
        .remove_slide_table_row(0, model_id, KeynoteTableRowDeletion::body(1))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    assert_eq!(
        editor.slide_tables(0).unwrap()[0].geometry,
        baseline_geometry
    );

    let before_error = editor.to_bytes().unwrap();
    assert!(
        editor
            .insert_slide_table_row(1, model_id, KeynoteTableRowInsertion::body(usize::MAX))
            .is_err()
    );
    assert!(
        editor
            .remove_slide_table_column(0, model_id, KeynoteTableColumnDeletion::body(usize::MAX))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_error);
}

#[test]
fn source_built_footer_formula_expands_and_contracts_with_body_rows() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Footer aggregate", 4, 3, position, size)
        .unwrap();
    let model_id = table.model_object_id;
    editor
        .set_slide_table_header_settings(
            0,
            model_id,
            KeynoteTableHeaderSettings {
                footer_rows: Some(KeynoteTableHeaderCount::ONE),
                ..Default::default()
            },
        )
        .unwrap();
    editor
        .set_slide_table_formula(
            0,
            model_id,
            3,
            1,
            KeynoteTableFormulaExpression::function(
                "SUM",
                [KeynoteTableFormulaExpression::range(
                    KeynoteTableFormulaCellReference::relative(1, 1),
                    KeynoteTableFormulaCellReference::relative(2, 1),
                )],
            ),
            KeynoteTableFormulaCachedValue::Number(3.0),
        )
        .unwrap();
    let mut package = editor.into_package();
    let engine = package.remove_entry("Index/CalculationEngine.iwa").unwrap();
    package
        .insert_entry("Index/CalculationEngine-42-2.iwa", engine)
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let baseline = editor.to_bytes().unwrap();

    editor
        .insert_slide_table_row(0, model_id, KeynoteTableRowInsertion::body(3))
        .unwrap();
    assert_eq!(
        editor
            .slide_table_formula(0, model_id, 4, 1)
            .unwrap()
            .as_deref(),
        Some("=SUM(B2:B4)")
    );
    editor
        .remove_slide_table_row(0, model_id, KeynoteTableRowDeletion::body(3))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn source_built_fixed_table_sections_roundtrip_full_axis_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Fixed sections", 4, 4, position, size)
        .unwrap();
    let model_id = table.model_object_id;
    editor
        .set_slide_table_header_settings(
            0,
            model_id,
            KeynoteTableHeaderSettings {
                header_rows: Some(KeynoteTableHeaderCount::ONE),
                header_columns: Some(KeynoteTableHeaderCount::ONE),
                footer_rows: Some(KeynoteTableHeaderCount::ONE),
                ..Default::default()
            },
        )
        .unwrap();
    let baseline = editor.to_bytes().unwrap();

    editor
        .insert_slide_table_row(0, model_id, KeynoteTableRowInsertion::header(1))
        .unwrap();
    editor
        .insert_slide_table_row(0, model_id, KeynoteTableRowInsertion::footer(0))
        .unwrap();
    editor
        .insert_slide_table_column(0, model_id, KeynoteTableColumnInsertion::header(1))
        .unwrap();
    let settings = editor.slide_table_header_settings(0, model_id).unwrap();
    assert_eq!(settings.header_row_count(), 2);
    assert_eq!(settings.footer_row_count(), 2);
    assert_eq!(settings.header_column_count(), 2);
    let table = editor.slide_table(0, model_id).unwrap();
    assert_eq!((table.info.rows, table.info.columns), (6, 5));

    editor
        .remove_slide_table_column(0, model_id, KeynoteTableColumnDeletion::header(1))
        .unwrap();
    editor
        .remove_slide_table_row(0, model_id, KeynoteTableRowDeletion::footer(0))
        .unwrap();
    editor
        .remove_slide_table_row(0, model_id, KeynoteTableRowDeletion::header(1))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_table_cell(
            0,
            model_id,
            0,
            0,
            KeynoteTableCellValue::Text("Header".to_owned()),
        )
        .unwrap();
    editor
        .set_slide_table_cell(
            0,
            model_id,
            1,
            1,
            KeynoteTableCellValue::Text("Body".to_owned()),
        )
        .unwrap();
    editor
        .set_slide_table_cell(
            0,
            model_id,
            3,
            2,
            KeynoteTableCellValue::Text("Footer".to_owned()),
        )
        .unwrap();
    editor
        .remove_slide_table_row(0, model_id, KeynoteTableRowDeletion::header(0))
        .unwrap();
    editor
        .remove_slide_table_row(0, model_id, KeynoteTableRowDeletion::footer(0))
        .unwrap();
    editor
        .remove_slide_table_column(0, model_id, KeynoteTableColumnDeletion::header(0))
        .unwrap();
    let settings = editor.slide_table_header_settings(0, model_id).unwrap();
    assert_eq!(settings.header_row_count(), 0);
    assert_eq!(settings.footer_row_count(), 0);
    assert_eq!(settings.header_column_count(), 0);
    let table = editor.slide_table(0, model_id).unwrap();
    assert_eq!((table.info.rows, table.info.columns), (2, 3));
    assert_eq!(
        table.get_cell(0, 0),
        Some(&KeynoteTableCellValue::Text("Body".to_owned()))
    );
    assert!(!table.cells.values().any(|value| matches!(
        value,
        KeynoteTableCellValue::Text(text) if text == "Header" || text == "Footer"
    )));
}

#[test]
fn source_built_table_roundtrips_title_settings_transactionally() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Forecast", 2, 2, position, size)
        .unwrap();
    let visible = KeynoteTableTitleSettings {
        visible: Some(true),
        outlined: Some(true),
    };
    let initially_hidden = KeynoteTableTitleSettings {
        visible: Some(false),
        outlined: None,
    };
    assert_eq!(
        editor
            .slide_table_title_settings(0, table.model_object_id)
            .unwrap(),
        initially_hidden
    );
    editor
        .set_slide_table_title_settings(0, table.model_object_id, visible)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_title_settings(0, table.model_object_id)
            .unwrap(),
        visible
    );
    let unchanged = reopened.to_bytes().unwrap();
    reopened
        .set_slide_table_title_settings(0, table.model_object_id, visible)
        .unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), unchanged);

    let explicit_hidden = KeynoteTableTitleSettings {
        visible: Some(false),
        outlined: Some(false),
    };
    reopened
        .set_slide_table_title_settings(0, table.model_object_id, explicit_hidden)
        .unwrap();
    assert_eq!(
        reopened
            .slide_table_title_settings(0, table.model_object_id)
            .unwrap(),
        explicit_hidden
    );
    reopened
        .set_slide_table_title_settings(
            0,
            table.model_object_id,
            KeynoteTableTitleSettings::default(),
        )
        .unwrap();
    assert_eq!(
        reopened
            .slide_table_title_settings(0, table.model_object_id)
            .unwrap(),
        KeynoteTableTitleSettings::default()
    );

    let before_error = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_slide_table_title_settings(1, table.model_object_id, visible)
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_error);
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
