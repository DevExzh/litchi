use super::*;
use crate::numbers::NumbersDocumentBuilder;
use crate::shapes::{RgbColorSpace, RgbaColor};
use crate::table_cell_conditional_highlight::{
    TableCellConditionalHighlightDate, TableCellConditionalHighlightDateOffset,
    TableCellConditionalHighlightDateOffsetDirection, TableCellConditionalHighlightDatePeriod,
    TableCellConditionalHighlightDatePeriodUnit, TableCellConditionalHighlightDateRange,
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

fn applied_rule(editor: &NumbersEditor, table_id: u64, row: usize, column: usize) -> Option<u32> {
    let location = locate_cell(&editor.package, table_id, row, column).unwrap();
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
fn equal_size_replacement_preserves_conditional_graph_identity() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(2, 2)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    let green = RgbaColor::new(0.1, 0.8, 0.2, 1.0, RgbColorSpace::Srgb).unwrap();
    let zero = TableCellConditionalHighlightNumber::new(0.0).unwrap();
    let initial = [
        rule(TableCellConditionalHighlightCondition::LessThan(zero), red),
        rule(
            TableCellConditionalHighlightCondition::GreaterThan(zero),
            green,
        ),
    ];
    editor
        .set_cell_conditional_highlighting(table_id, 1, 1, &initial)
        .unwrap();
    let before = info_in_package(&editor.package, table_id, 1, 1)
        .unwrap()
        .unwrap();
    let location = locate_cell(&editor.package, table_id, 1, 1).unwrap();
    let child_identifiers = conditional_style_rule_identifiers(
        &editor.package,
        &location.object_locations,
        before.style_set_object_id,
    )
    .unwrap();

    let replacement = [
        rule(
            TableCellConditionalHighlightCondition::LessThanOrEqualTo(zero),
            green,
        ),
        rule(
            TableCellConditionalHighlightCondition::GreaterThanOrEqualTo(zero),
            red,
        ),
    ];
    editor
        .set_cell_conditional_highlighting(table_id, 1, 1, &replacement)
        .unwrap();

    let after = info_in_package(&editor.package, table_id, 1, 1)
        .unwrap()
        .unwrap();
    let location = locate_cell(&editor.package, table_id, 1, 1).unwrap();
    assert_eq!(after.list_identifier, before.list_identifier);
    assert_eq!(after.style_set_object_id, before.style_set_object_id);
    assert_eq!(
        conditional_style_rule_identifiers(
            &editor.package,
            &location.object_locations,
            after.style_set_object_id,
        )
        .unwrap(),
        child_identifiers
    );
    assert_eq!(
        editor
            .cell_conditional_highlight_rules(table_id, 1, 1)
            .unwrap(),
        Some(replacement.to_vec())
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

    assert_eq!(applied_rule(&editor, table_id, 1, 1), Some(0));
    assert_eq!(applied_rule(&editor, table_id, 1, 2), Some(1));

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
fn relative_date_predicates_round_trip_apply_and_observe_day_boundaries() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(3, 5)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let context = current_date_context();
    let today = context.apple_seconds;
    let midday = SECONDS_PER_DAY / 2.0;
    for (column, value) in [
        (1, today + midday),
        (2, today - SECONDS_PER_DAY + midday),
        (3, today + SECONDS_PER_DAY + midday),
    ] {
        editor
            .set_cell(table_id, 1, column, CellValue::Date(value))
            .unwrap();
    }
    editor
        .set_cell(table_id, 1, 4, CellValue::Number(today + midday))
        .unwrap();
    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    let green = RgbaColor::new(0.1, 0.8, 0.2, 1.0, RgbColorSpace::Srgb).unwrap();
    let blue = RgbaColor::new(0.1, 0.2, 0.9, 1.0, RgbColorSpace::Srgb).unwrap();
    let rules = [
        rule(TableCellConditionalHighlightCondition::DateIsToday, red),
        rule(
            TableCellConditionalHighlightCondition::DateIsYesterday,
            green,
        ),
        rule(TableCellConditionalHighlightCondition::DateIsTomorrow, blue),
    ];

    for column in 1..=4 {
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
    assert_eq!(applied_rule(&editor, table_id, 1, 1), Some(0));
    assert_eq!(applied_rule(&editor, table_id, 1, 2), Some(1));
    assert_eq!(applied_rule(&editor, table_id, 1, 3), Some(2));
    assert_eq!(
        applied_rule(&editor, table_id, 1, 4),
        Some(CONDITIONAL_STYLE_NO_APPLIED_RULE)
    );

    let info = info_in_package(&editor.package, table_id, 1, 1)
        .unwrap()
        .unwrap();
    let location = locate_cell(&editor.package, table_id, 1, 1).unwrap();
    let archive = editor
        .package
        .archive(&location.object_locations[&info.style_set_object_id])
        .unwrap();
    let object = archive.object(info.style_set_object_id).unwrap();
    let set = object
        .messages
        .iter()
        .find_map(|message| {
            (message.type_ == 6_010)
                .then(|| tst::ConditionalStyleSetArchive::decode(message.data.as_slice()).unwrap())
        })
        .unwrap();
    let current = &set.rules.as_ref().unwrap().rule;
    for (index, (kind, node_count)) in [
        (RelativeDatePredicateKind::Today, 9),
        (RelativeDatePredicateKind::Yesterday, 9),
        (RelativeDatePredicateKind::Tomorrow, 11),
    ]
    .into_iter()
    .enumerate()
    {
        let predicate = current[index].predicate.as_ref().unwrap();
        let prepivot = &set.rules_prepivot[index].predicate;
        assert_eq!(predicate.predicate_type, kind.native_value());
        assert_eq!(prepivot.predicate_type, kind.native_value());
        assert_eq!(prepivot.param_index0, PREDICATE_CELL_ARGUMENT_INDEX);
        assert_eq!(prepivot.param_index1, PREDICATE_UNUSED_ARGUMENT_INDEX);
        assert_eq!(prepivot.param_index2, PREDICATE_UNUSED_ARGUMENT_INDEX);
        assert_eq!(
            predicate.formula.as_ref().unwrap().ast_node_array.ast_node,
            prepivot.formula.ast_node_array.ast_node
        );
        assert_eq!(
            predicate
                .formula
                .as_ref()
                .unwrap()
                .ast_node_array
                .ast_node
                .len(),
            node_count
        );
    }

    for (condition, start) in [
        (
            TableCellConditionalHighlightCondition::DateIsYesterday,
            today - SECONDS_PER_DAY,
        ),
        (TableCellConditionalHighlightCondition::DateIsToday, today),
        (
            TableCellConditionalHighlightCondition::DateIsTomorrow,
            today + SECONDS_PER_DAY,
        ),
    ] {
        assert!(condition_matches_at(
            &condition,
            &ConditionalCellValue::Date(start),
            Some(context)
        ));
        assert!(condition_matches_at(
            &condition,
            &ConditionalCellValue::Date(start + SECONDS_PER_DAY - 1.0),
            Some(context)
        ));
        assert!(!condition_matches_at(
            &condition,
            &ConditionalCellValue::Date(start - 1.0),
            Some(context)
        ));
        assert!(!condition_matches_at(
            &condition,
            &ConditionalCellValue::Date(start + SECONDS_PER_DAY),
            Some(context)
        ));
    }
}

#[test]
fn date_period_predicates_round_trip_all_units_and_use_calendar_boundaries() {
    use TableCellConditionalHighlightDateOffsetDirection as Direction;
    use TableCellConditionalHighlightDatePeriodUnit as Unit;

    let units = [
        Unit::Days,
        Unit::Weeks,
        Unit::Months,
        Unit::Quarters,
        Unit::Years,
    ];
    let mut conditions = Vec::with_capacity(20);
    for unit in units {
        let period = TableCellConditionalHighlightDatePeriod::new(2, unit).unwrap();
        conditions.extend([
            TableCellConditionalHighlightCondition::DateIsInNext(period),
            TableCellConditionalHighlightCondition::DateIsInLast(period),
            TableCellConditionalHighlightCondition::DateIsOffsetFromToday(
                TableCellConditionalHighlightDateOffset::new(period, Direction::Ago),
            ),
            TableCellConditionalHighlightCondition::DateIsOffsetFromToday(
                TableCellConditionalHighlightDateOffset::new(period, Direction::FromNow),
            ),
        ]);
    }
    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(2, 3)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    for (chunk_index, condition_chunk) in conditions.chunks(10).enumerate() {
        let column = chunk_index + 1;
        let rules: Vec<_> = condition_chunk
            .iter()
            .cloned()
            .map(|condition| rule(condition, red))
            .collect();
        editor
            .set_cell_conditional_highlighting(table_id, 1, column, &rules)
            .unwrap();
        assert_eq!(
            editor
                .cell_conditional_highlight_rules(table_id, 1, column)
                .unwrap(),
            Some(rules)
        );

        let info = info_in_package(&editor.package, table_id, 1, column)
            .unwrap()
            .unwrap();
        let location = locate_cell(&editor.package, table_id, 1, column).unwrap();
        let archive = editor
            .package
            .archive(&location.object_locations[&info.style_set_object_id])
            .unwrap();
        let set = archive
            .object(info.style_set_object_id)
            .unwrap()
            .messages
            .iter()
            .find_map(|message| {
                (message.type_ == 6_010).then(|| {
                    tst::ConditionalStyleSetArchive::decode(message.data.as_slice()).unwrap()
                })
            })
            .unwrap();
        for (index, condition) in condition_chunk.iter().enumerate() {
            let kind = NativePredicateKind::from_condition(condition);
            let NativePredicateKind::DatePeriod(period_kind) = kind else {
                panic!("expected date-period predicate");
            };
            let current = set.rules.as_ref().unwrap().rule[index]
                .predicate
                .as_ref()
                .unwrap();
            let prepivot = &set.rules_prepivot[index].predicate;
            assert_eq!(current.predicate_type, period_kind.native_value());
            assert_eq!(prepivot.predicate_type, period_kind.native_value());
            assert_eq!(
                prepivot.param_index1,
                formula::date_period_quantity_node_index(period_kind, condition).unwrap()
            );
            assert_eq!(
                current.formula.as_ref().unwrap().ast_node_array.ast_node,
                prepivot.formula.ast_node_array.ast_node
            );
        }
    }

    let today = NaiveDate::from_ymd_opt(2024, 1, 31).unwrap();
    let context = ConditionalDateContext {
        today,
        apple_seconds: date_to_apple_seconds(today),
    };
    let month = TableCellConditionalHighlightDatePeriod::new(1, Unit::Months).unwrap();
    let leap_day = date_to_apple_seconds(NaiveDate::from_ymd_opt(2024, 2, 29).unwrap());
    let exact = TableCellConditionalHighlightCondition::DateIsOffsetFromToday(
        TableCellConditionalHighlightDateOffset::new(month, Direction::FromNow),
    );
    assert!(condition_matches_at(
        &exact,
        &ConditionalCellValue::Date(leap_day),
        Some(context)
    ));
    assert!(!condition_matches_at(
        &exact,
        &ConditionalCellValue::Date(leap_day + SECONDS_PER_DAY),
        Some(context)
    ));
}

#[test]
fn volatile_date_predicates_register_native_dependency_owners_and_tiles() {
    use TableCellConditionalHighlightDatePeriodUnit as Unit;

    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    let period = TableCellConditionalHighlightDatePeriod::new(2, Unit::Days).unwrap();
    let volatile_rule = rule(
        TableCellConditionalHighlightCondition::DateIsInNext(period),
        red,
    );
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(2, 3)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    for column in 1..=2 {
        editor
            .set_cell_conditional_highlighting(
                table_id,
                1,
                column,
                std::slice::from_ref(&volatile_rule),
            )
            .unwrap();
    }

    let entry = editor
        .package
        .calculation_engine_entry_name()
        .unwrap()
        .unwrap()
        .to_owned();
    let archive = editor.package.archive(&entry).unwrap();
    let owners = archive
        .objects
        .iter()
        .filter_map(|object| {
            object.messages.iter().find_map(|message| {
                (message.type_ == 4_008).then(|| {
                    (
                        object.archive_info.identifier.unwrap(),
                        tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())
                            .unwrap(),
                    )
                })
            })
        })
        .collect::<Vec<_>>();
    assert_eq!(owners.len(), 13);
    let mut internal_kinds = owners
        .iter()
        .filter_map(|(_, owner)| {
            (owner.owner_kind != Some(1))
                .then_some((owner.internal_formula_owner_id, owner.owner_kind.unwrap()))
        })
        .collect::<Vec<_>>();
    internal_kinds.sort_unstable();
    assert_eq!(
        internal_kinds
            .iter()
            .map(|(_, kind)| *kind)
            .collect::<Vec<_>>(),
        vec![7, 200, 4, 11, 3, 10, 5, 6, 35, 12, 8, 9]
    );
    let engine = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find_map(|message| {
            (message.type_ == 4_000)
                .then(|| tsce::CalculationEngineArchive::decode(message.data.as_slice()).unwrap())
        })
        .unwrap();
    assert_eq!(engine.dependency_tracker.number_of_formulas, Some(9));
    let root_kinds = engine
        .dependency_tracker
        .formula_owner_dependencies
        .iter()
        .map(|reference| {
            owners
                .iter()
                .find(|(identifier, _)| *identifier == reference.identifier)
                .unwrap()
                .1
                .owner_kind
                .unwrap()
        })
        .collect::<Vec<_>>();
    assert_eq!(
        root_kinds,
        vec![9, 8, 12, 35, 6, 5, 10, 3, 11, 4, 200, 1, 7]
    );
    let tiled_kinds = owners
        .iter()
        .filter_map(|(_, owner)| {
            (!owner
                .tiled_cell_dependencies
                .as_ref()
                .unwrap()
                .cell_record_tiles
                .is_empty())
            .then_some(owner.owner_kind.unwrap())
        })
        .collect::<HashSet<_>>();
    assert_eq!(tiled_kinds, HashSet::from([8, 35, 3, 200]));
    let conditional = owners
        .iter()
        .map(|(_, owner)| owner)
        .find(|owner| owner.owner_kind == Some(3))
        .unwrap();
    let records = &conditional.cell_dependencies.as_ref().unwrap().cell_record;
    assert_eq!(
        records
            .iter()
            .map(|record| (record.row, record.column))
            .collect::<Vec<_>>(),
        vec![(1, 1), (1, 2)]
    );
    let tile_id = conditional
        .tiled_cell_dependencies
        .as_ref()
        .unwrap()
        .cell_record_tiles[0]
        .identifier;
    let tile = archive
        .object(tile_id)
        .unwrap()
        .messages
        .iter()
        .find_map(|message| {
            (message.type_ == 4_009)
                .then(|| tsce::CellRecordTileArchive::decode(message.data.as_slice()).unwrap())
        })
        .unwrap();
    assert_eq!(tile.cell_records, *records);

    editor
        .clear_cell_conditional_highlighting(table_id, 1, 1)
        .unwrap();
    let archive = editor.package.archive(&entry).unwrap();
    let engine = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find_map(|message| {
            (message.type_ == 4_000)
                .then(|| tsce::CalculationEngineArchive::decode(message.data.as_slice()).unwrap())
        })
        .unwrap();
    assert_eq!(engine.dependency_tracker.number_of_formulas, Some(8));
    let conditional = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find_map(|message| {
            (message.type_ == 4_008)
                .then(|| {
                    tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice()).unwrap()
                })
                .filter(|owner| owner.owner_kind == Some(3))
        })
        .unwrap();
    assert_eq!(
        conditional
            .cell_dependencies
            .as_ref()
            .unwrap()
            .cell_record
            .iter()
            .map(|record| (record.row, record.column))
            .collect::<Vec<_>>(),
        vec![(1, 2)]
    );
    assert!(dependencies::removal::contains_coordinate(
        conditional
            .volatile_dependencies
            .as_ref()
            .unwrap()
            .volatile_time_cells
            .as_ref()
            .unwrap(),
        1,
        2
    ));
    assert!(!dependencies::removal::contains_coordinate(
        conditional
            .volatile_dependencies
            .as_ref()
            .unwrap()
            .volatile_time_cells
            .as_ref()
            .unwrap(),
        1,
        1
    ));

    let exact = TableCellConditionalHighlightDate::from_ymd(2026, 7, 28).unwrap();
    let fixed_rule = rule(TableCellConditionalHighlightCondition::DateIs(exact), red);
    editor
        .set_cell_conditional_highlighting(table_id, 1, 2, std::slice::from_ref(&fixed_rule))
        .unwrap();
    let archive = editor.package.archive(&entry).unwrap();
    let engine = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find_map(|message| {
            (message.type_ == 4_000)
                .then(|| tsce::CalculationEngineArchive::decode(message.data.as_slice()).unwrap())
        })
        .unwrap();
    assert_eq!(engine.dependency_tracker.number_of_formulas, Some(7));
    let conditional = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find_map(|message| {
            (message.type_ == 4_008)
                .then(|| {
                    tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice()).unwrap()
                })
                .filter(|owner| owner.owner_kind == Some(3))
        })
        .unwrap();
    assert!(
        conditional
            .cell_dependencies
            .as_ref()
            .unwrap()
            .cell_record
            .is_empty()
    );
    assert!(
        conditional
            .volatile_dependencies
            .as_ref()
            .unwrap()
            .volatile_time_cells
            .as_ref()
            .unwrap()
            .column_entries
            .is_empty()
    );
    assert!(
        conditional
            .uuid_references
            .as_ref()
            .unwrap()
            .table_refs
            .is_empty()
    );
    let tile_id = conditional
        .tiled_cell_dependencies
        .as_ref()
        .unwrap()
        .cell_record_tiles[0]
        .identifier;
    let tile = archive
        .object(tile_id)
        .unwrap()
        .messages
        .iter()
        .find_map(|message| {
            (message.type_ == 4_009)
                .then(|| tsce::CellRecordTileArchive::decode(message.data.as_slice()).unwrap())
        })
        .unwrap();
    assert!(tile.cell_records.is_empty());
}

#[test]
fn fixed_date_predicates_round_trip_apply_and_preserve_native_graphs() {
    let exact = TableCellConditionalHighlightDate::from_ymd(2026, 7, 27).unwrap();
    let lower = TableCellConditionalHighlightDate::from_ymd(2026, 7, 26).unwrap();
    let upper = TableCellConditionalHighlightDate::from_ymd(2026, 7, 28).unwrap();
    let range = TableCellConditionalHighlightDateRange::new(lower, upper).unwrap();
    let conditions = [
        TableCellConditionalHighlightCondition::DateIs(exact),
        TableCellConditionalHighlightCondition::DateIsBefore(exact),
        TableCellConditionalHighlightCondition::DateIsAfter(exact),
        TableCellConditionalHighlightCondition::DateIsBetween(range),
    ];
    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    let rules = conditions.clone().map(|condition| rule(condition, red));
    let midday = SECONDS_PER_DAY / 2.0;
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(3, 6)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    for (column, value) in [
        (1, exact.apple_seconds() + midday),
        (2, lower.apple_seconds() + midday),
        (3, upper.apple_seconds() + midday),
    ] {
        editor
            .set_cell(table_id, 1, column, CellValue::Date(value))
            .unwrap();
        editor
            .set_cell_conditional_highlighting(table_id, 1, column, &rules)
            .unwrap();
    }
    editor
        .set_cell(
            table_id,
            1,
            4,
            CellValue::Date(exact.apple_seconds() + midday),
        )
        .unwrap();
    editor
        .set_cell_conditional_highlighting(table_id, 1, 4, &rules[3..])
        .unwrap();
    editor
        .set_cell(
            table_id,
            1,
            5,
            CellValue::Number(exact.apple_seconds() + midday),
        )
        .unwrap();
    editor
        .set_cell_conditional_highlighting(table_id, 1, 5, &rules)
        .unwrap();

    assert_eq!(applied_rule(&editor, table_id, 1, 1), Some(0));
    assert_eq!(applied_rule(&editor, table_id, 1, 2), Some(1));
    assert_eq!(applied_rule(&editor, table_id, 1, 3), Some(2));
    assert_eq!(applied_rule(&editor, table_id, 1, 4), Some(0));
    assert_eq!(
        applied_rule(&editor, table_id, 1, 5),
        Some(CONDITIONAL_STYLE_NO_APPLIED_RULE)
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
    let archive = editor
        .package
        .archive(&location.object_locations[&info.style_set_object_id])
        .unwrap();
    let object = archive.object(info.style_set_object_id).unwrap();
    let set = object
        .messages
        .iter()
        .find_map(|message| {
            (message.type_ == 6_010)
                .then(|| tst::ConditionalStyleSetArchive::decode(message.data.as_slice()).unwrap())
        })
        .unwrap();
    let current = &set.rules.as_ref().unwrap().rule;
    for (index, (kind, node_count, indexes)) in [
        (
            FixedDatePredicateKind::Equal,
            16,
            (
                PREDICATE_DATE_EQUALITY_CELL_ARGUMENT_INDEX,
                PREDICATE_DATE_EQUALITY_ARGUMENT_INDEX,
                PREDICATE_UNUSED_ARGUMENT_INDEX,
            ),
        ),
        (
            FixedDatePredicateKind::Before,
            7,
            (
                PREDICATE_CELL_ARGUMENT_INDEX,
                PREDICATE_DATE_ARGUMENT_INDEX,
                PREDICATE_UNUSED_ARGUMENT_INDEX,
            ),
        ),
        (
            FixedDatePredicateKind::After,
            5,
            (
                PREDICATE_CELL_ARGUMENT_INDEX,
                PREDICATE_DATE_ARGUMENT_INDEX,
                PREDICATE_UNUSED_ARGUMENT_INDEX,
            ),
        ),
        (
            FixedDatePredicateKind::Between,
            26,
            (
                PREDICATE_RANGE_CELL_ARGUMENT_INDEX,
                PREDICATE_RANGE_LOWER_ARGUMENT_INDEX,
                PREDICATE_RANGE_UPPER_ARGUMENT_INDEX,
            ),
        ),
    ]
    .into_iter()
    .enumerate()
    {
        let predicate = current[index].predicate.as_ref().unwrap();
        let prepivot = &set.rules_prepivot[index].predicate;
        assert_eq!(predicate.predicate_type, kind.native_value());
        assert_eq!(
            predicate.param_value1.as_ref().unwrap().arg_type,
            PREDICATE_ARGUMENT_DATE
        );
        assert_eq!(
            predicate.param_value2.as_ref().unwrap().arg_type,
            if kind.is_range() {
                PREDICATE_ARGUMENT_DATE
            } else {
                PREDICATE_ARGUMENT_NONE
            }
        );
        assert_eq!(prepivot.predicate_type, kind.native_value());
        assert_eq!(
            (
                prepivot.param_index0,
                prepivot.param_index1,
                prepivot.param_index2,
            ),
            indexes
        );
        assert_eq!(
            predicate.formula.as_ref().unwrap().ast_node_array.ast_node,
            prepivot.formula.ast_node_array.ast_node
        );
        assert_eq!(
            predicate
                .formula
                .as_ref()
                .unwrap()
                .ast_node_array
                .ast_node
                .len(),
            node_count
        );
    }

    let exact_start = exact.apple_seconds();
    assert!(condition_matches(
        &conditions[0],
        &ConditionalCellValue::Date(exact_start)
    ));
    assert!(condition_matches(
        &conditions[0],
        &ConditionalCellValue::Date(exact_start + SECONDS_PER_DAY - 1.0)
    ));
    assert!(!condition_matches(
        &conditions[0],
        &ConditionalCellValue::Date(exact_start + SECONDS_PER_DAY)
    ));
    assert!(condition_matches(
        &conditions[1],
        &ConditionalCellValue::Date(exact_start - 1.0)
    ));
    assert!(!condition_matches(
        &conditions[1],
        &ConditionalCellValue::Date(exact_start)
    ));
    assert!(!condition_matches(
        &conditions[2],
        &ConditionalCellValue::Date(exact_start + SECONDS_PER_DAY - 1.0)
    ));
    assert!(condition_matches(
        &conditions[2],
        &ConditionalCellValue::Date(exact_start + SECONDS_PER_DAY)
    ));
    assert!(condition_matches(
        &conditions[3],
        &ConditionalCellValue::Date(upper.apple_seconds() + SECONDS_PER_DAY - 1.0)
    ));
    assert!(!condition_matches(
        &conditions[3],
        &ConditionalCellValue::Date(upper.apple_seconds() + SECONDS_PER_DAY)
    ));
}

#[test]
fn numeric_sign_predicates_round_trip_and_exclude_zero_and_non_numbers() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(3, 4)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    for (column, value) in [(1, 3.0), (2, -3.0), (3, 0.0)] {
        editor
            .set_cell(table_id, 1, column, CellValue::Number(value))
            .unwrap();
    }
    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    let green = RgbaColor::new(0.1, 0.8, 0.2, 1.0, RgbColorSpace::Srgb).unwrap();
    let rules = [
        rule(
            TableCellConditionalHighlightCondition::NumberIsPositive,
            red,
        ),
        rule(
            TableCellConditionalHighlightCondition::NumberIsNegative,
            green,
        ),
    ];
    for column in 1..=3 {
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

    assert_eq!(applied_rule(&editor, table_id, 1, 1), Some(0));
    assert_eq!(applied_rule(&editor, table_id, 1, 2), Some(1));
    assert_eq!(
        applied_rule(&editor, table_id, 1, 3),
        Some(CONDITIONAL_STYLE_NO_APPLIED_RULE)
    );
    let info = info_in_package(&editor.package, table_id, 1, 1)
        .unwrap()
        .unwrap();
    let location = locate_cell(&editor.package, table_id, 1, 1).unwrap();
    let archive_name = &location.object_locations[&info.style_set_object_id];
    let archive = editor.package.archive(archive_name).unwrap();
    let object = archive.object(info.style_set_object_id).unwrap();
    let set = object
        .messages
        .iter()
        .find_map(|message| {
            (message.type_ == 6_010)
                .then(|| tst::ConditionalStyleSetArchive::decode(message.data.as_slice()).unwrap())
        })
        .unwrap();
    let current = set.rules.unwrap().rule;
    for (index, (expected_current, expected_prepivot)) in [
        (
            NumericSignPredicateKind::IsPositive.native_value(),
            NumericPredicateKind::GreaterThan.native_value(),
        ),
        (
            NumericSignPredicateKind::IsNegative.native_value(),
            NumericPredicateKind::LessThan.native_value(),
        ),
    ]
    .into_iter()
    .enumerate()
    {
        let current = current[index].predicate.as_ref().unwrap();
        let prepivot = &set.rules_prepivot[index].predicate;
        assert_eq!(current.predicate_type, expected_current);
        assert_eq!(
            current
                .formula
                .as_ref()
                .unwrap()
                .ast_node_array
                .ast_node
                .len(),
            6
        );
        assert_eq!(prepivot.predicate_type, expected_prepivot);
        assert_eq!(prepivot.param_index0, PREDICATE_CELL_ARGUMENT_INDEX);
        assert_eq!(prepivot.param_index1, PREDICATE_NUMBER_ARGUMENT_INDEX);
        assert_eq!(prepivot.param_index2, PREDICATE_UNUSED_ARGUMENT_INDEX);
        assert_eq!(prepivot.formula.ast_node_array.ast_node.len(), 3);
    }
    assert!(!condition_matches(
        &TableCellConditionalHighlightCondition::NumberIsPositive,
        &ConditionalCellValue::Other,
    ));
    assert!(!condition_matches(
        &TableCellConditionalHighlightCondition::NumberIsNegative,
        &ConditionalCellValue::Number(0.0),
    ));
}

#[test]
fn boolean_predicates_round_trip_and_match_only_the_exact_boolean() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(3, 5)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    for (column, value) in [
        (1, CellValue::Boolean(true)),
        (2, CellValue::Boolean(false)),
        (3, CellValue::Number(1.0)),
        (4, CellValue::Text("TRUE".to_owned())),
    ] {
        editor.set_cell(table_id, 1, column, value).unwrap();
    }
    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    let green = RgbaColor::new(0.1, 0.8, 0.2, 1.0, RgbColorSpace::Srgb).unwrap();
    let rules = [
        rule(TableCellConditionalHighlightCondition::BooleanIsTrue, red),
        rule(
            TableCellConditionalHighlightCondition::BooleanIsFalse,
            green,
        ),
    ];
    for column in 1..=4 {
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

    assert_eq!(applied_rule(&editor, table_id, 1, 1), Some(0));
    assert_eq!(applied_rule(&editor, table_id, 1, 2), Some(1));
    for column in [3, 4] {
        assert_eq!(
            applied_rule(&editor, table_id, 1, column),
            Some(CONDITIONAL_STYLE_NO_APPLIED_RULE)
        );
    }

    let info = info_in_package(&editor.package, table_id, 1, 1)
        .unwrap()
        .unwrap();
    let location = locate_cell(&editor.package, table_id, 1, 1).unwrap();
    let archive = editor
        .package
        .archive(&location.object_locations[&info.style_set_object_id])
        .unwrap();
    let object = archive.object(info.style_set_object_id).unwrap();
    let set = object
        .messages
        .iter()
        .find_map(|message| {
            (message.type_ == 6_010)
                .then(|| tst::ConditionalStyleSetArchive::decode(message.data.as_slice()).unwrap())
        })
        .unwrap();
    let current = set.rules.unwrap().rule;
    for (index, kind) in [BooleanPredicateKind::IsTrue, BooleanPredicateKind::IsFalse]
        .into_iter()
        .enumerate()
    {
        let predicate = current[index].predicate.as_ref().unwrap();
        let prepivot = &set.rules_prepivot[index].predicate;
        assert_eq!(predicate.predicate_type, kind.native_value());
        assert_eq!(
            predicate
                .formula
                .as_ref()
                .unwrap()
                .ast_node_array
                .ast_node
                .len(),
            [13, 14][index]
        );
        let nodes = &predicate.formula.as_ref().unwrap().ast_node_array.ast_node;
        assert_eq!(
            nodes[1].ast_function_node_index,
            Some(VALUE_TYPE_FUNCTION_INDEX)
        );
        assert_eq!(
            nodes[2].ast_number_node_number,
            Some(BOOLEAN_VALUE_TYPE_CODE)
        );
        assert_eq!(
            prepivot.predicate_type,
            NumericPredicateKind::EqualTo.native_value()
        );
        assert_eq!(prepivot.param_index0, PREDICATE_CELL_ARGUMENT_INDEX);
        assert_eq!(prepivot.param_index1, PREDICATE_NUMBER_ARGUMENT_INDEX);
        assert_eq!(prepivot.param_index2, PREDICATE_UNUSED_ARGUMENT_INDEX);
        assert_eq!(
            prepivot.formula.ast_node_array.ast_node[1].ast_boolean_node_boolean,
            Some(kind.value())
        );
    }

    assert!(!condition_matches(
        &TableCellConditionalHighlightCondition::BooleanIsTrue,
        &ConditionalCellValue::Number(1.0),
    ));
    assert!(!condition_matches(
        &TableCellConditionalHighlightCondition::BooleanIsFalse,
        &ConditionalCellValue::Text("FALSE".to_owned()),
    ));
}

#[test]
fn checkbox_predicates_round_trip_and_require_native_checkbox_format() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(3, 5)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    for (column, value) in [(1, true), (2, false), (3, true), (4, false)] {
        editor
            .set_cell(table_id, 1, column, CellValue::Boolean(value))
            .unwrap();
    }
    for column in [1, 2] {
        editor
            .set_table_cell_checkbox_format(
                table_id,
                1,
                column,
                crate::table_cell_data_format::TableCellCheckboxFormat,
            )
            .unwrap();
    }
    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    let green = RgbaColor::new(0.1, 0.8, 0.2, 1.0, RgbColorSpace::Srgb).unwrap();
    let rules = [
        rule(
            TableCellConditionalHighlightCondition::CheckboxIsChecked,
            red,
        ),
        rule(
            TableCellConditionalHighlightCondition::CheckboxIsNotChecked,
            green,
        ),
    ];
    for column in 1..=4 {
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

    assert_eq!(applied_rule(&editor, table_id, 1, 1), Some(0));
    assert_eq!(applied_rule(&editor, table_id, 1, 2), Some(1));
    for column in [3, 4] {
        assert_eq!(
            applied_rule(&editor, table_id, 1, column),
            Some(CONDITIONAL_STYLE_NO_APPLIED_RULE)
        );
    }

    let info = info_in_package(&editor.package, table_id, 1, 1)
        .unwrap()
        .unwrap();
    let location = locate_cell(&editor.package, table_id, 1, 1).unwrap();
    let archive = editor
        .package
        .archive(&location.object_locations[&info.style_set_object_id])
        .unwrap();
    let object = archive.object(info.style_set_object_id).unwrap();
    let set = object
        .messages
        .iter()
        .find_map(|message| {
            (message.type_ == 6_010)
                .then(|| tst::ConditionalStyleSetArchive::decode(message.data.as_slice()).unwrap())
        })
        .unwrap();
    let current = set.rules.unwrap().rule;
    for (index, kind) in [
        CheckboxPredicateKind::IsChecked,
        CheckboxPredicateKind::IsNotChecked,
    ]
    .into_iter()
    .enumerate()
    {
        let predicate = current[index].predicate.as_ref().unwrap();
        let prepivot = &set.rules_prepivot[index].predicate;
        assert_eq!(predicate.predicate_type, kind.native_value());
        assert_eq!(
            predicate
                .formula
                .as_ref()
                .unwrap()
                .ast_node_array
                .ast_node
                .len(),
            [13, 14][index]
        );
        let nodes = &predicate.formula.as_ref().unwrap().ast_node_array.ast_node;
        let format_function_index = if kind.is_checked() { 2 } else { 3 };
        assert_eq!(
            nodes[format_function_index].ast_function_node_index,
            Some(CELL_DATA_FORMAT_FUNCTION_INDEX)
        );
        assert_eq!(
            nodes[format_function_index + 1].ast_number_node_number,
            Some(CHECKBOX_DATA_FORMAT_CODE)
        );
        assert_eq!(
            prepivot.predicate_type,
            NumericPredicateKind::EqualTo.native_value()
        );
        assert_eq!(
            prepivot.formula.ast_node_array.ast_node[1].ast_boolean_node_boolean,
            Some(kind.is_checked())
        );
    }

    assert!(!condition_matches(
        &TableCellConditionalHighlightCondition::CheckboxIsChecked,
        &ConditionalCellValue::Boolean(true),
    ));
    assert!(condition_matches(
        &TableCellConditionalHighlightCondition::BooleanIsTrue,
        &ConditionalCellValue::Checkbox(true),
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
