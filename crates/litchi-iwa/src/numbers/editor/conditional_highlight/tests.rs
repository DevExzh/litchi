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
fn blank_predicates_round_trip_and_apply_to_empty_and_populated_cells() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(3, 3)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    editor
        .set_cell(table_id, 1, 2, CellValue::Text("occupied".to_owned()))
        .unwrap();
    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    let green = RgbaColor::new(0.1, 0.8, 0.2, 1.0, RgbColorSpace::Srgb).unwrap();
    let rules = [
        rule(TableCellConditionalHighlightCondition::CellIsBlank, red),
        rule(
            TableCellConditionalHighlightCondition::CellIsNotBlank,
            green,
        ),
    ];

    for column in [1, 2] {
        editor
            .set_cell_conditional_highlighting(table_id, 1, column, &rules)
            .unwrap();
        assert_eq!(
            editor
                .cell_conditional_highlight_rules(table_id, 1, column)
                .unwrap(),
            Some(rules.to_vec())
        );
    }

    let applied_rule = |editor: &NumbersEditor, column| {
        let location = locate_cell(&editor.package, table_id, 1, column).unwrap();
        let cell = read_tile_cell(
            &editor.package,
            &location.tile_archive,
            location.tile_id,
            location.tile_row,
            column,
        )
        .unwrap()
        .unwrap();
        BncCell::parse(&cell)
            .unwrap()
            .conditional_style_applied_rule()
    };
    assert_eq!(applied_rule(&editor, 1), Some(0));
    assert_eq!(applied_rule(&editor, 2), Some(1));

    assert!(condition_matches(
        &TableCellConditionalHighlightCondition::CellIsBlank,
        &ConditionalCellValue::Blank,
    ));
    assert!(!condition_matches(
        &TableCellConditionalHighlightCondition::CellIsNotBlank,
        &ConditionalCellValue::Blank,
    ));
    assert!(condition_matches(
        &TableCellConditionalHighlightCondition::CellIsNotBlank,
        &ConditionalCellValue::Other,
    ));
}

#[test]
fn blank_predicates_reject_noncanonical_formula_graphs() {
    let mut editor = NumbersDocumentBuilder::new().build().unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    editor
        .set_cell_conditional_highlighting(
            table_id,
            0,
            0,
            &[rule(
                TableCellConditionalHighlightCondition::CellIsBlank,
                red,
            )],
        )
        .unwrap();
    let info = info_in_package(&editor.package, table_id, 0, 0)
        .unwrap()
        .unwrap();
    let location = locate_cell(&editor.package, table_id, 0, 0).unwrap();
    let archive_name = location.object_locations[&info.style_set_object_id].clone();
    editor
        .package
        .update_archive(&archive_name, |archive| {
            let object = archive
                .object_mut(info.style_set_object_id)
                .ok_or_else(|| Error::InvalidFormat("style set missing".to_owned()))?;
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == 6_010)
                .ok_or_else(|| Error::InvalidFormat("style-set payload missing".to_owned()))?;
            let mut set = tst::ConditionalStyleSetArchive::decode(
                object.messages[message_index].data.as_slice(),
            )?;
            set.rules
                .as_mut()
                .and_then(|rules| rules.rule[0].predicate.as_mut())
                .and_then(|predicate| predicate.formula.as_mut())
                .and_then(|formula| formula.ast_node_array.ast_node.get_mut(1))
                .ok_or_else(|| Error::InvalidFormat("current formula missing".to_owned()))?
                .ast_function_node_index = Some(IS_ERROR_FUNCTION_INDEX);
            set.rules_prepivot[0]
                .predicate
                .formula
                .ast_node_array
                .ast_node[1]
                .ast_function_node_index = Some(IS_ERROR_FUNCTION_INDEX);
            let message_type = object.messages[message_index].type_;
            object.replace_message(
                message_index,
                crate::archive::RawMessage {
                    type_: message_type,
                    data: set.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    assert!(
        editor
            .cell_conditional_highlight_rules(table_id, 0, 0)
            .is_err()
    );
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
            TableCellConditionalHighlightCondition::TextDoesNotStartWith(text("dairy")),
            "Dairy Milk",
        ),
        (
            TableCellConditionalHighlightCondition::TextEndsWith(text("grain")),
            "Dairy",
        ),
        (
            TableCellConditionalHighlightCondition::TextDoesNotEndWith(text("rice")),
            "Organic Rice",
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

    let info = info_in_package(&editor.package, table_id, 1, 1)
        .unwrap()
        .unwrap();
    let location = locate_cell(&editor.package, table_id, 1, 1).unwrap();
    let archive_name = location.object_locations[&info.style_set_object_id].clone();
    editor
        .package
        .update_archive(&archive_name, |archive| {
            let object = archive
                .object_mut(info.style_set_object_id)
                .ok_or_else(|| Error::InvalidFormat("style set missing".to_owned()))?;
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == 6_010)
                .ok_or_else(|| Error::InvalidFormat("style-set payload missing".to_owned()))?;
            let mut set = tst::ConditionalStyleSetArchive::decode(
                object.messages[message_index].data.as_slice(),
            )?;
            set.rules_prepivot.clear();
            let message_type = object.messages[message_index].type_;
            object.replace_message(
                message_index,
                crate::archive::RawMessage {
                    type_: message_type,
                    data: set.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    assert_eq!(
        editor
            .cell_conditional_highlight_rules(table_id, 1, 1)
            .unwrap(),
        Some(rules.to_vec())
    );
}
