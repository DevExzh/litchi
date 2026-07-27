//! Canonical formula graphs for relative-day conditional-highlight predicates.

use super::*;

const SECONDS_PER_DAY: f64 = 86_400.0;
const DURATION_UNIT: i32 = 3;
const DURATION_STYLE: u32 = 1;
const DURATION_LARGEST_UNIT: u32 = 4;
const DURATION_SMALLEST_UNIT: u32 = 16;
const ONE_DAY: u32 = 1;
const TWO_DAYS: u32 = 2;

pub(super) fn nodes(
    kind: RelativeDatePredicateKind,
    formula_owner_uuid: &tsp::Uuid,
) -> Vec<tsce::ast_node_array_archive::AstNodeArchive> {
    use tsce::ast_node_array_archive::AstNodeType;

    match kind {
        RelativeDatePredicateKind::Today => vec![
            linked_cell_node(formula_owner_uuid),
            today_node(),
            operator_node(AstNodeType::GreaterThanOrEqualToNode),
            linked_cell_node(formula_owner_uuid),
            today_node(),
            duration_node(ONE_DAY),
            operator_node(AstNodeType::AdditionNode),
            operator_node(AstNodeType::LessThanNode),
            function_node(LOGICAL_AND_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
        ],
        RelativeDatePredicateKind::Yesterday => vec![
            linked_cell_node(formula_owner_uuid),
            today_node(),
            duration_node(ONE_DAY),
            operator_node(AstNodeType::SubtractionNode),
            operator_node(AstNodeType::GreaterThanOrEqualToNode),
            linked_cell_node(formula_owner_uuid),
            today_node(),
            operator_node(AstNodeType::LessThanNode),
            function_node(LOGICAL_AND_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
        ],
        RelativeDatePredicateKind::Tomorrow => vec![
            linked_cell_node(formula_owner_uuid),
            today_node(),
            duration_node(ONE_DAY),
            operator_node(AstNodeType::AdditionNode),
            operator_node(AstNodeType::GreaterThanOrEqualToNode),
            linked_cell_node(formula_owner_uuid),
            today_node(),
            duration_node(TWO_DAYS),
            operator_node(AstNodeType::AdditionNode),
            operator_node(AstNodeType::LessThanNode),
            function_node(LOGICAL_AND_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
        ],
    }
}

pub(super) fn validate(
    formula: &tsce::FormulaArchive,
    kind: RelativeDatePredicateKind,
) -> Result<()> {
    let expected = nodes(kind, &tsp::Uuid { upper: 0, lower: 0 });
    let actual = &formula.ast_node_array.ast_node;
    if actual.len() != expected.len()
        || actual
            .iter()
            .zip(expected.iter())
            .any(|(actual, expected)| {
                if node_matches(
                    expected,
                    tsce::ast_node_array_archive::AstNodeType::LinkedCellRefNode,
                ) {
                    !node_matches(
                        actual,
                        tsce::ast_node_array_archive::AstNodeType::LinkedCellRefNode,
                    )
                } else {
                    actual != expected
                }
            })
    {
        return Err(invalid_formula());
    }
    Ok(())
}

fn today_node() -> tsce::ast_node_array_archive::AstNodeArchive {
    function_node(TODAY_FUNCTION_INDEX, ZERO_FUNCTION_ARGUMENT_COUNT)
}

fn duration_node(days: u32) -> tsce::ast_node_array_archive::AstNodeArchive {
    use tsce::ast_node_array_archive::AstNodeType;

    tsce::ast_node_array_archive::AstNodeArchive {
        ast_node_type: AstNodeType::DurationNode as i32,
        ast_duration_node_unit_num: Some(f64::from(days) * SECONDS_PER_DAY),
        ast_duration_node_unit: Some(DURATION_UNIT),
        ast_duration_node_style: Some(DURATION_STYLE),
        ast_duration_node_duration_unit_largest: Some(DURATION_LARGEST_UNIT),
        ast_duration_node_duration_unit_smallest: Some(DURATION_SMALLEST_UNIT),
        ast_duration_node_use_automatic_units: Some(true),
        ..Default::default()
    }
}
