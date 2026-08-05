use super::*;
use crate::numbers::NumbersDocumentBuilder;
use crate::shapes::{RgbColorSpace, RgbaColor};
use litchi_iwa_common::table::cell::conditional_highlight::{
    Date, Offset,
    OffsetDirection, Period,
    PeriodUnit, DateRange,
    Number, Range,
    Style, Text,
};

fn rule(
    condition: Condition,
    color: RgbaColor,
) -> Rule {
    Rule::new(
        condition,
        Style::with_fill(color),
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

fn share_conditional_highlight(
    editor: &mut NumbersEditor,
    table_id: u64,
    row: usize,
    source_column: usize,
    target_column: usize,
) -> TableCellConditionalHighlightInfo {
    let shared = info_in_package(&editor.package, table_id, row, source_column)
        .unwrap()
        .unwrap();
    let target = locate_cell(&editor.package, table_id, row, target_column).unwrap();
    let (resolved, entry) =
        resolve_entry(&editor.package, &target, shared.list_identifier).unwrap();
    increment_table_data_list_entry(
        &mut editor.package,
        &target.object_locations,
        &resolved,
        &entry,
        tst::table_data_list::ListType::ConditionalStyle,
    )
    .unwrap();
    update_cell(
        &mut editor.package,
        &target,
        row,
        target_column,
        Some(shared.list_identifier),
        Some(CONDITIONAL_STYLE_NO_APPLIED_RULE),
    )
    .unwrap();
    shared
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
    let zero = Number::new(0.0).unwrap();
    let hundred = Number::new(100.0).unwrap();
    let range = Range::new(zero, hundred).unwrap();
    let initial = [
        rule(Condition::LessThan(zero), red),
        rule(
            Condition::GreaterThanOrEqualTo(hundred),
            green,
        ),
        rule(
            Condition::Between(range),
            green,
        ),
        rule(
            Condition::NotBetween(range),
            red,
        ),
    ];

    let created = editor
        .set_cell_conditional_highlighting(table_id, 1, 1, &initial)
        .unwrap();
    assert_eq!(created.table_id, table_id);
    assert_eq!((created.row, created.column), (1, 1));
    assert_eq!(created.rule_count, initial.len() as u32);
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
    let graph = info_in_package(&editor.package, table_id, 1, 1)
        .unwrap()
        .unwrap();
    let location = locate_cell(&editor.package, table_id, 1, 1).unwrap();
    let children = replace::conditional_style_rule_identifiers(
        &editor.package,
        &location.object_locations,
        graph.style_set_object_id,
    )
    .unwrap()
    .into_iter()
    .flat_map(|identifiers| [identifiers.text_style, identifiers.cell_style])
    .collect::<Vec<_>>();
    let graph_identifiers = children
        .iter()
        .copied()
        .chain(std::iter::once(graph.style_set_object_id))
        .collect::<Vec<_>>();
    let graph_components = graph_identifiers
        .iter()
        .filter_map(|identifier| {
            location
                .object_locations
                .get(identifier)
                .map(|archive_name| {
                    component_identifier_for_entry(&editor.package, archive_name)
                        .unwrap()
                        .unwrap()
                })
        })
        .collect::<HashSet<_>>();
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
    for identifier in graph_identifiers {
        assert!(editor.package.iwa_entry_names().all(|name| {
            editor
                .package
                .archive(name)
                .unwrap()
                .object(identifier)
                .is_none()
        }));
        assert!(graph_components.iter().all(|component| {
            component_uuid_identifiers(&editor.package, *component)
                .unwrap()
                .is_none_or(|registered| !registered.contains(&identifier))
        }));
    }
}

#[test]
fn shared_conditional_child_rejects_deletion_transactionally() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(2, 2)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    let zero = Number::new(0.0).unwrap();
    editor
        .set_cell_conditional_highlighting(
            table_id,
            1,
            1,
            &[rule(
                Condition::LessThan(zero),
                red,
            )],
        )
        .unwrap();
    let graph = info_in_package(&editor.package, table_id, 1, 1)
        .unwrap()
        .unwrap();
    let location = locate_cell(&editor.package, table_id, 1, 1).unwrap();
    let child = replace::conditional_style_rule_identifiers(
        &editor.package,
        &location.object_locations,
        graph.style_set_object_id,
    )
    .unwrap()[0]
        .text_style;
    let model_archive = location.object_locations[&table_id].clone();
    editor
        .package
        .update_archive(&model_archive, |archive| {
            let model = archive
                .object_mut(table_id)
                .ok_or_else(|| Error::InvalidFormat("table model is missing".to_owned()))?;
            model.archive_info.message_infos[0]
                .object_references
                .push(child);
            Ok(())
        })
        .unwrap();
    let before = editor.to_bytes().unwrap();

    let error = editor
        .clear_cell_conditional_highlighting(table_id, 1, 1)
        .unwrap_err();

    assert!(error.to_string().contains("shared by another object"));
    assert_eq!(editor.to_bytes().unwrap(), before);
    assert!(
        editor
            .cell_conditional_highlighting(table_id, 1, 1)
            .unwrap()
            .is_some()
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
    let zero = Number::new(0.0).unwrap();
    let initial = [
        rule(Condition::LessThan(zero), red),
        rule(
            Condition::GreaterThan(zero),
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
    let child_identifiers = replace::conditional_style_rule_identifiers(
        &editor.package,
        &location.object_locations,
        before.style_set_object_id,
    )
    .unwrap();

    let replacement = [
        rule(
            Condition::LessThanOrEqualTo(zero),
            green,
        ),
        rule(
            Condition::GreaterThanOrEqualTo(zero),
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
        replace::conditional_style_rule_identifiers(
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
fn variable_size_replacement_retains_and_reclaims_conditional_children() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(2, 2)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    let green = RgbaColor::new(0.1, 0.8, 0.2, 1.0, RgbColorSpace::Srgb).unwrap();
    let zero = Number::new(0.0).unwrap();
    let one = Number::new(1.0).unwrap();
    let initial = [rule(
        Condition::LessThan(zero),
        red,
    )];
    editor
        .set_cell_conditional_highlighting(table_id, 1, 1, &initial)
        .unwrap();
    let before = info_in_package(&editor.package, table_id, 1, 1)
        .unwrap()
        .unwrap();
    let location = locate_cell(&editor.package, table_id, 1, 1).unwrap();
    let initial_children = replace::conditional_style_rule_identifiers(
        &editor.package,
        &location.object_locations,
        before.style_set_object_id,
    )
    .unwrap();

    let grown = [
        rule(
            Condition::LessThanOrEqualTo(zero),
            green,
        ),
        rule(
            Condition::GreaterThanOrEqualTo(one),
            red,
        ),
        rule(Condition::EqualTo(one), green),
    ];
    editor
        .set_cell_conditional_highlighting(table_id, 1, 1, &grown)
        .unwrap();
    let grown_info = info_in_package(&editor.package, table_id, 1, 1)
        .unwrap()
        .unwrap();
    let grown_location = locate_cell(&editor.package, table_id, 1, 1).unwrap();
    let grown_children = replace::conditional_style_rule_identifiers(
        &editor.package,
        &grown_location.object_locations,
        grown_info.style_set_object_id,
    )
    .unwrap();
    assert_eq!(grown_info.list_identifier, before.list_identifier);
    assert_eq!(grown_info.style_set_object_id, before.style_set_object_id);
    assert_eq!(grown_children[0], initial_children[0]);
    assert_eq!(grown_children.len(), grown.len());

    let removed = grown_children[1..]
        .iter()
        .flat_map(|identifiers| [identifiers.text_style, identifiers.cell_style])
        .collect::<Vec<_>>();
    editor
        .set_cell_conditional_highlighting(table_id, 1, 1, &grown[..1])
        .unwrap();
    let shrunk_info = info_in_package(&editor.package, table_id, 1, 1)
        .unwrap()
        .unwrap();
    let shrunk_location = locate_cell(&editor.package, table_id, 1, 1).unwrap();
    let shrunk_children = replace::conditional_style_rule_identifiers(
        &editor.package,
        &shrunk_location.object_locations,
        shrunk_info.style_set_object_id,
    )
    .unwrap();
    assert_eq!(shrunk_info.list_identifier, before.list_identifier);
    assert_eq!(shrunk_info.style_set_object_id, before.style_set_object_id);
    assert_eq!(shrunk_children, initial_children);
    for identifier in removed {
        assert!(editor.package.iwa_entry_names().all(|name| {
            editor
                .package
                .archive(name)
                .unwrap()
                .object(identifier)
                .is_none()
        }));
    }
    assert_eq!(
        editor
            .cell_conditional_highlight_rules(table_id, 1, 1)
            .unwrap(),
        Some(grown[..1].to_vec())
    );
}

#[test]
fn shared_conditional_graph_replacement_is_copy_on_write() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(2, 3)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    let green = RgbaColor::new(0.1, 0.8, 0.2, 1.0, RgbColorSpace::Srgb).unwrap();
    let zero = Number::new(0.0).unwrap();
    let initial = [rule(
        Condition::LessThan(zero),
        red,
    )];
    editor
        .set_cell_conditional_highlighting(table_id, 1, 1, &initial)
        .unwrap();
    let shared = share_conditional_highlight(&mut editor, table_id, 1, 1, 2);

    let replacement = [rule(
        Condition::GreaterThanOrEqualTo(zero),
        green,
    )];
    editor
        .set_cell_conditional_highlighting(table_id, 1, 1, &replacement)
        .unwrap();

    let edited = info_in_package(&editor.package, table_id, 1, 1)
        .unwrap()
        .unwrap();
    let untouched = info_in_package(&editor.package, table_id, 1, 2)
        .unwrap()
        .unwrap();
    assert_ne!(edited.list_identifier, shared.list_identifier);
    assert_ne!(edited.style_set_object_id, shared.style_set_object_id);
    assert_eq!(untouched.list_identifier, shared.list_identifier);
    assert_eq!(untouched.style_set_object_id, shared.style_set_object_id);
    assert_eq!(
        editor
            .cell_conditional_highlight_rules(table_id, 1, 1)
            .unwrap(),
        Some(replacement.to_vec())
    );
    assert_eq!(
        editor
            .cell_conditional_highlight_rules(table_id, 1, 2)
            .unwrap(),
        Some(initial.to_vec())
    );
    for (info, expected_refcount) in [(edited, 1), (untouched, 1)] {
        let location = locate_cell(&editor.package, table_id, info.row, info.column).unwrap();
        let (_resolved, entry) =
            resolve_entry(&editor.package, &location, info.list_identifier).unwrap();
        assert_eq!(entry.entry.refcount, expected_refcount);
    }
}

#[test]
fn shared_conditional_copy_on_write_rolls_back_after_late_failure() {
    let mut editor = NumbersDocumentBuilder::new()
        .table_dimensions(2, 3)
        .build()
        .unwrap();
    let table_id = editor.tables().unwrap()[0].object_id;
    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    let zero = Number::new(0.0).unwrap();
    let rules = [rule(
        Condition::LessThan(zero),
        red,
    )];
    editor
        .set_cell_conditional_highlighting(table_id, 1, 1, &rules)
        .unwrap();
    let shared = share_conditional_highlight(&mut editor, table_id, 1, 1, 2);
    let location = locate_cell(&editor.package, table_id, 1, 1).unwrap();
    let (resolved, _entry) =
        resolve_entry(&editor.package, &location, shared.list_identifier).unwrap();
    editor
        .package
        .update_archive(&resolved.table_archive, |archive| {
            let object = archive.object_mut(resolved.table_id).ok_or_else(|| {
                Error::InvalidFormat("conditional-style table missing".to_owned())
            })?;
            let message_index = table_data_list_message_index(
                object,
                tst::table_data_list::ListType::ConditionalStyle,
            )
            .ok_or_else(|| Error::InvalidFormat("conditional-style payload missing".to_owned()))?;
            let mut list = TableDataList::decode(object.messages[message_index].data.as_slice())?;
            list.next_list_id = u32::MAX;
            let message_type = object.messages[message_index].type_;
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data: list.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let next_object_before = next_object_identifier(&editor.package).unwrap();

    assert!(
        editor
            .set_cell_conditional_highlighting(table_id, 1, 1, &rules)
            .is_err()
    );

    assert_eq!(
        next_object_identifier(&editor.package).unwrap(),
        next_object_before
    );
    let location = locate_cell(&editor.package, table_id, 1, 1).unwrap();
    let (_resolved, entry) =
        resolve_entry(&editor.package, &location, shared.list_identifier).unwrap();
    assert_eq!(entry.entry.refcount, 2);
    assert_eq!(
        info_in_package(&editor.package, table_id, 1, 1)
            .unwrap()
            .unwrap()
            .style_set_object_id,
        shared.style_set_object_id
    );
    assert_eq!(
        info_in_package(&editor.package, table_id, 1, 2)
            .unwrap()
            .unwrap()
            .style_set_object_id,
        shared.style_set_object_id
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
    let lower = Number::new(3.0).unwrap();
    let upper = Number::new(7.0).unwrap();
    let range = Range::new(lower, upper).unwrap();
    let between = Condition::Between(range);
    let not_between = Condition::NotBetween(range);

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
        rule(Condition::CellIsBlank, red),
        rule(
            Condition::CellIsNotBlank,
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
        &Condition::CellIsBlank,
        &ConditionalCellValue::Blank,
    ));
    assert!(!condition_matches(
        &Condition::CellIsNotBlank,
        &ConditionalCellValue::Blank,
    ));
    assert!(condition_matches(
        &Condition::CellIsNotBlank,
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
        rule(Condition::DateIsToday, red),
        rule(
            Condition::DateIsYesterday,
            green,
        ),
        rule(Condition::DateIsTomorrow, blue),
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
            Condition::DateIsYesterday,
            today - SECONDS_PER_DAY,
        ),
        (Condition::DateIsToday, today),
        (
            Condition::DateIsTomorrow,
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
    use OffsetDirection as Direction;
    use PeriodUnit as Unit;

    let units = [
        Unit::Days,
        Unit::Weeks,
        Unit::Months,
        Unit::Quarters,
        Unit::Years,
    ];
    let mut conditions = Vec::with_capacity(20);
    for unit in units {
        let period = Period::new(2, unit).unwrap();
        conditions.extend([
            Condition::DateIsInNext(period),
            Condition::DateIsInLast(period),
            Condition::DateIsOffsetFromToday(
                Offset::new(period, Direction::Ago),
            ),
            Condition::DateIsOffsetFromToday(
                Offset::new(period, Direction::FromNow),
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
    let month = Period::new(1, Unit::Months).unwrap();
    let leap_day = date_to_apple_seconds(NaiveDate::from_ymd_opt(2024, 2, 29).unwrap());
    let exact = Condition::DateIsOffsetFromToday(
        Offset::new(month, Direction::FromNow),
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
    use PeriodUnit as Unit;

    let red = RgbaColor::new(0.9, 0.1, 0.1, 1.0, RgbColorSpace::Srgb).unwrap();
    let period = Period::new(2, Unit::Days).unwrap();
    let volatile_rule = rule(
        Condition::DateIsInNext(period),
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

    let exact = Date::from_ymd(2026, 7, 28).unwrap();
    let fixed_rule = rule(Condition::DateIs(exact), red);
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
    let exact = Date::from_ymd(2026, 7, 27).unwrap();
    let lower = Date::from_ymd(2026, 7, 26).unwrap();
    let upper = Date::from_ymd(2026, 7, 28).unwrap();
    let range = DateRange::new(lower, upper).unwrap();
    let conditions = [
        Condition::DateIs(exact),
        Condition::DateIsBefore(exact),
        Condition::DateIsAfter(exact),
        Condition::DateIsBetween(range),
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
            Condition::NumberIsPositive,
            red,
        ),
        rule(
            Condition::NumberIsNegative,
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
        &Condition::NumberIsPositive,
        &ConditionalCellValue::Other,
    ));
    assert!(!condition_matches(
        &Condition::NumberIsNegative,
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
        rule(Condition::BooleanIsTrue, red),
        rule(
            Condition::BooleanIsFalse,
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
        &Condition::BooleanIsTrue,
        &ConditionalCellValue::Number(1.0),
    ));
    assert!(!condition_matches(
        &Condition::BooleanIsFalse,
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
            Condition::CheckboxIsChecked,
            red,
        ),
        rule(
            Condition::CheckboxIsNotChecked,
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
        &Condition::CheckboxIsChecked,
        &ConditionalCellValue::Boolean(true),
    ));
    assert!(condition_matches(
        &Condition::BooleanIsTrue,
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
                Condition::CellIsBlank,
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
    let text = |value| Text::new(value).unwrap();
    let condition_cases = [
        (
            Condition::TextEqualTo(text("organic grain")),
            "Dairy",
        ),
        (
            Condition::TextNotEqualTo(text("dairy")),
            "Dairy",
        ),
        (
            Condition::TextStartsWith(text("ORGANIC")),
            "Dairy",
        ),
        (
            Condition::TextDoesNotStartWith(text("dairy")),
            "Dairy Milk",
        ),
        (
            Condition::TextEndsWith(text("grain")),
            "Dairy",
        ),
        (
            Condition::TextDoesNotEndWith(text("rice")),
            "Organic Rice",
        ),
        (
            Condition::TextContains(text("NIC GR")),
            "Dairy",
        ),
        (
            Condition::TextDoesNotContain(text("rice")),
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
    editor
        .set_cell_conditional_highlighting(table_id, 1, 1, &rules[..1])
        .unwrap();
    let replaced = info_in_package(&editor.package, table_id, 1, 1)
        .unwrap()
        .unwrap();
    assert_eq!(replaced.list_identifier, info.list_identifier);
    assert_eq!(replaced.style_set_object_id, info.style_set_object_id);
    assert_eq!(
        editor
            .cell_conditional_highlight_rules(table_id, 1, 1)
            .unwrap(),
        Some(rules[..1].to_vec())
    );
}
