use litchi_odf::{
    ConditionalCellStyle, ConditionalCellStyleRule, FormulaNamespace, MutableSpreadsheet,
    Spreadsheet, SpreadsheetBuilder,
};

fn rule(condition: &str, target: &str) -> ConditionalCellStyleRule {
    ConditionalCellStyleRule::new(condition, target)
        .with_formula_namespace(FormulaNamespace {
            prefix: "of".to_string(),
            uri: "urn:oasis:names:tc:opendocument:xmlns:of:1.2".to_string(),
        })
        .with_base_cell_address("Sheet1.A1")
}

#[test]
fn builder_and_mutable_conditional_styles_have_stable_packaged_roundtrips() {
    let mut builder = SpreadsheetBuilder::new();
    builder.add_sheet("Sheet1").unwrap();
    builder.add_common_table_cell_style("Red").unwrap();
    builder.add_common_table_cell_style("Blue").unwrap();
    let original = ConditionalCellStyle::new("ce1", vec![rule("of:cell-content()>0", "Red")])
        .with_parent_style_name("Base");
    builder
        .create_conditional_cell_style(original.clone())
        .unwrap();

    let spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(
        spreadsheet.conditional_cell_styles(),
        std::slice::from_ref(&original)
    );

    let mut mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
    let replacement = ConditionalCellStyle::new("ce1", vec![rule("of:cell-content()<=10", "Blue")]);
    assert_eq!(
        mutable
            .replace_conditional_cell_style(replacement.clone())
            .unwrap(),
        original
    );
    let reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.conditional_cell_styles(),
        std::slice::from_ref(&replacement)
    );

    let mut mutable = MutableSpreadsheet::from_spreadsheet(reopened).unwrap();
    assert_eq!(
        mutable.remove_conditional_cell_style("ce1").unwrap(),
        Some(replacement)
    );
    let reopened = Spreadsheet::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert!(reopened.conditional_cell_styles().is_empty());
}

#[test]
fn validation_failures_are_atomic_and_conditions_remain_inert() {
    let mut builder = SpreadsheetBuilder::new();
    builder.add_common_table_cell_style("Red").unwrap();
    let valid = ConditionalCellStyle::new("ce1", vec![rule("of:is-true-formula([.A1])", "Red")]);
    builder
        .create_conditional_cell_style(valid.clone())
        .unwrap();

    let invalid = ConditionalCellStyle::new(
        "ce2",
        vec![ConditionalCellStyleRule::new("bad:condition()", "Missing")],
    );
    assert!(builder.create_conditional_cell_style(invalid).is_err());
    assert_eq!(builder.conditional_cell_styles(), [valid]);
    assert!(
        builder
            .create_conditional_cell_style(ConditionalCellStyle::new("", vec![]))
            .is_err()
    );
    assert_eq!(builder.conditional_cell_styles().len(), 1);
}
