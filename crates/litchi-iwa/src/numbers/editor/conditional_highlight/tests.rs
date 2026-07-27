use super::*;
use crate::numbers::NumbersDocumentBuilder;
use crate::shapes::{RgbColorSpace, RgbaColor};
use crate::table_cell_conditional_highlight::{
    TableCellConditionalHighlightNumber, TableCellConditionalHighlightRange,
    TableCellConditionalHighlightStyle, TableCellConditionalHighlightText,
};

fn rule(
    condition: TableCellConditionalHighlightCondition,
    color: RgbaColor,
) -> TableCellConditionalHighlightRule {
    TableCellConditionalHighlightRule::new(
        condition,
        TableCellConditionalHighlightStyle::with_fill(color),
    )
}

#[test]
fn scratch_document_conditional_highlights_create_replace_and_delete() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(4, 3)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    let green = RgbaColor::new(0.1, 0.8, 0.2, 1.0, RgbColorSpace::Srgb).unwrap();
    let zero = TableCellConditionalHighlightNumber::new(0.0).unwrap();
    let hundred = TableCellConditionalHighlightNumber::new(100.0).unwrap();
    let range = TableCellConditionalHighlightRange::new(zero, hundred).unwrap();
    let initial = [
        rule(TableCellConditionalHighlightCondition::LessThan(zero), red),
        rule(
            TableCellConditionalHighlightCondition::GreaterThanOrEqualTo(hundred),
            green,
        ),
        rule(
            TableCellConditionalHighlightCondition::Between(range),
            green,
        ),
        rule(
            TableCellConditionalHighlightCondition::NotBetween(range),
            red,
        ),
    ];

    editor
        .set_cell_conditional_highlighting(table_id, 1, 1, &initial)
        .unwrap();
    assert_eq!(
        editor
            .cell_conditional_highlighting(table_id, 1, 1)
            .unwrap()
            .unwrap()
            .rule_count,
        4
    );
    assert_eq!(
        editor
            .cell_conditional_highlight_rules(table_id, 1, 1)
            .unwrap()
            .unwrap(),
        initial
    );

    editor
        .set_cell_conditional_highlighting(table_id, 1, 1, &initial[..1])
        .unwrap();
    assert_eq!(
        editor
            .cell_conditional_highlighting(table_id, 1, 1)
            .unwrap()
            .unwrap()
            .rule_count,
        1
    );
    assert_eq!(
        editor
            .cell_conditional_highlight_rules(table_id, 1, 1)
            .unwrap()
            .unwrap(),
        initial[..1]
    );
    editor
        .clear_cell_conditional_highlighting(table_id, 1, 1)
        .unwrap();
    assert!(
        editor
            .cell_conditional_highlighting(table_id, 1, 1)
            .unwrap()
            .is_none()
    );
    assert!(
        editor
            .cell_conditional_highlight_rules(table_id, 1, 1)
            .unwrap()
            .is_none()
    );
}

#[test]
fn empty_or_excessive_rule_sets_are_rejected_transactionally() {
    let mut editor = NumbersDocumentBuilder::new().build().unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    assert!(
        editor
            .set_cell_conditional_highlighting(table_id, 0, 0, &[])
            .is_err()
    );
    assert!(
        editor
            .cell_conditional_highlighting(table_id, 0, 0)
            .unwrap()
            .is_none()
    );
}

#[test]
fn range_conditions_use_inclusive_between_and_strictly_outside_not_between() {
    let lower = TableCellConditionalHighlightNumber::new(3.0).unwrap();
    let upper = TableCellConditionalHighlightNumber::new(7.0).unwrap();
    let range = TableCellConditionalHighlightRange::new(lower, upper).unwrap();
    let between = TableCellConditionalHighlightCondition::Between(range);
    let not_between = TableCellConditionalHighlightCondition::NotBetween(range);

    assert!(condition_matches(
        &between,
        &ConditionalCellValue::Number(3.0)
    ));
    assert!(condition_matches(
        &between,
        &ConditionalCellValue::Number(5.0)
    ));
    assert!(condition_matches(
        &between,
        &ConditionalCellValue::Number(7.0)
    ));
    assert!(!condition_matches(
        &between,
        &ConditionalCellValue::Number(2.0)
    ));
    assert!(!condition_matches(
        &not_between,
        &ConditionalCellValue::Number(3.0)
    ));
    assert!(!condition_matches(
        &not_between,
        &ConditionalCellValue::Number(7.0)
    ));
    assert!(condition_matches(
        &not_between,
        &ConditionalCellValue::Number(2.0)
    ));
    assert!(condition_matches(
        &not_between,
        &ConditionalCellValue::Number(8.0)
    ));
}

#[test]
fn text_predicates_are_case_insensitive_and_round_trip_from_scratch() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(3, 3)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    editor
        .set_cell(table_id, 1, 1, CellValue::Text("Organic Grain".to_owned()))
        .unwrap();
    let text = |value| TableCellConditionalHighlightText::new(value).unwrap();
    let condition_cases = [
        (
            TableCellConditionalHighlightCondition::TextEqualTo(text("organic grain")),
            "Dairy",
        ),
        (
            TableCellConditionalHighlightCondition::TextNotEqualTo(text("dairy")),
            "Dairy",
        ),
        (
            TableCellConditionalHighlightCondition::TextStartsWith(text("ORGANIC")),
            "Dairy",
        ),
        (
            TableCellConditionalHighlightCondition::TextEndsWith(text("grain")),
            "Dairy",
        ),
        (
            TableCellConditionalHighlightCondition::TextContains(text("NIC GR")),
            "Dairy",
        ),
        (
            TableCellConditionalHighlightCondition::TextDoesNotContain(text("rice")),
            "Organic Rice",
        ),
    ];
    for (condition, nonmatching) in &condition_cases {
        assert!(condition_matches(
            condition,
            &ConditionalCellValue::Text("Organic Grain".to_owned())
        ));
        assert!(!condition_matches(
            condition,
            &ConditionalCellValue::Text((*nonmatching).to_owned())
        ));
    }
    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    let rules = condition_cases.map(|(condition, _nonmatching)| rule(condition, red));

    editor
        .set_cell_conditional_highlighting(table_id, 1, 1, &rules)
        .unwrap();
    let location = locate_cell(&editor.package, table_id, 1, 1).unwrap();
    let cell = read_tile_cell(
        &editor.package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        1,
    )
    .unwrap()
    .unwrap();
    assert_eq!(
        BncCell::parse(&cell)
            .unwrap()
            .conditional_style_applied_rule(),
        Some(0)
    );
    assert_eq!(
        editor
            .cell_conditional_highlight_rules(table_id, 1, 1)
            .unwrap(),
        Some(rules.to_vec())
    );
}
