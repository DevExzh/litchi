use litchi_odf::{
    CellStyleProtection, ConditionalCellStyle, ConditionalCellStyleRule, MutableSpreadsheet,
    Spreadsheet, SpreadsheetBuilder, TableCellProtectionStyle,
};

#[test]
fn every_protection_value_round_trips_through_builder_and_mutable_packages() {
    let values = [
        CellStyleProtection::None,
        CellStyleProtection::Protected,
        CellStyleProtection::FormulaHidden,
        CellStyleProtection::ProtectedFormulaHidden,
        CellStyleProtection::HiddenAndProtected,
    ];
    let mut builder = SpreadsheetBuilder::new();
    builder.add_sheet("Sheet1").unwrap();
    builder.add_common_table_cell_style("Base").unwrap();
    for (index, value) in values.into_iter().enumerate() {
        builder
            .create_table_cell_protection_style(
                TableCellProtectionStyle::new(format!("p{index}"), value)
                    .with_parent_style_name("Base"),
            )
            .unwrap();
    }
    builder.set_cell_style_name(0, 0, "p1").unwrap();
    let mut spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(spreadsheet.table_cell_protection_styles().len(), values.len());
    let cell = spreadsheet.sheets().unwrap()[0].rows[0].cells[0].clone();
    assert_eq!(
        spreadsheet.cell_style_protection(&cell).unwrap(),
        Some(CellStyleProtection::Protected)
    );

    let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
    let old = mutable
        .replace_table_cell_protection_style(
            TableCellProtectionStyle::new("p1", CellStyleProtection::HiddenAndProtected)
                .with_parent_style_name("Base"),
        )
        .unwrap();
    assert_eq!(old.protection, CellStyleProtection::Protected);
    let mut reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    let cell = reopened.sheets().unwrap()[0].rows[0].cells[0].clone();
    assert_eq!(
        reopened.cell_style_protection(&cell).unwrap(),
        Some(CellStyleProtection::HiddenAndProtected)
    );
}

#[test]
fn combined_conditional_and_protection_style_survives_replace_and_remove() {
    let mut builder = SpreadsheetBuilder::new();
    builder.add_common_table_cell_style("Red").unwrap();
    let conditional = ConditionalCellStyle::new(
        "combo",
        vec![ConditionalCellStyleRule::new("cell-content()>0", "Red")],
    );
    builder
        .create_conditional_cell_style(conditional.clone())
        .unwrap();
    builder
        .create_table_cell_protection_style(TableCellProtectionStyle::new(
            "combo",
            CellStyleProtection::Protected,
        ))
        .unwrap();
    let spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(
        spreadsheet.conditional_cell_styles(),
        std::slice::from_ref(&conditional)
    );
    assert_eq!(spreadsheet.table_cell_protection_styles().len(), 1);

    let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
    mutable
        .replace_table_cell_protection_style(TableCellProtectionStyle::new(
            "combo",
            CellStyleProtection::FormulaHidden,
        ))
        .unwrap();
    let reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.conditional_cell_styles(), [conditional]);
    assert_eq!(
        reopened.table_cell_protection_styles()[0].protection,
        CellStyleProtection::FormulaHidden
    );

    let mut mutable = MutableSpreadsheet::from_spreadsheet(reopened).unwrap();
    mutable
        .remove_table_cell_protection_style("combo")
        .unwrap();
    let reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.conditional_cell_styles().len(), 1);
    assert!(reopened.table_cell_protection_styles().is_empty());
}

#[test]
fn malformed_names_parents_duplicates_and_failed_replacements_are_atomic() {
    let mut builder = SpreadsheetBuilder::new();
    builder.add_common_table_cell_style("Base").unwrap();
    let valid = TableCellProtectionStyle::new("p", CellStyleProtection::Protected)
        .with_parent_style_name("Base");
    builder
        .create_table_cell_protection_style(valid.clone())
        .unwrap();
    assert!(builder
        .create_table_cell_protection_style(TableCellProtectionStyle::new(
            "",
            CellStyleProtection::None,
        ))
        .is_err());
    assert!(builder
        .replace_table_cell_protection_style(
            TableCellProtectionStyle::new("p", CellStyleProtection::FormulaHidden)
                .with_parent_style_name("Missing"),
        )
        .is_err());
    assert_eq!(builder.table_cell_protection_styles(), [valid]);
}
