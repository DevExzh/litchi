//! Canonical formula graphs for calendar-date conditional-highlight predicates.

use super::*;

const SECONDS_PER_DAY: f64 = 86_400.0;
const DURATION_UNIT: i32 = 3;
const DURATION_STYLE: u32 = 1;
const DURATION_LARGEST_UNIT: u32 = 4;
const DURATION_SMALLEST_UNIT: u32 = 16;
const ONE_DAY: u32 = 1;
const TWO_DAYS: u32 = 2;

pub(super) fn relative_nodes(
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

pub(super) fn validate_relative(
    formula: &tsce::FormulaArchive,
    kind: RelativeDatePredicateKind,
) -> Result<()> {
    let expected = relative_nodes(kind, &tsp::Uuid { upper: 0, lower: 0 });
    validate_nodes(formula, &expected)
}

pub(super) fn fixed_nodes(
    kind: FixedDatePredicateKind,
    lower: f64,
    upper: Option<f64>,
    formula_owner_uuid: &tsp::Uuid,
) -> Vec<tsce::ast_node_array_archive::AstNodeArchive> {
    use tsce::ast_node_array_archive::AstNodeType;

    match kind {
        FixedDatePredicateKind::Equal => {
            let mut nodes = Vec::with_capacity(16);
            for component_function in [
                DATE_YEAR_FUNCTION_INDEX,
                DATE_MONTH_FUNCTION_INDEX,
                DATE_DAY_FUNCTION_INDEX,
            ] {
                nodes.extend([
                    date_node(lower),
                    function_node(component_function, UNARY_FUNCTION_ARGUMENT_COUNT),
                    linked_cell_node(formula_owner_uuid),
                    function_node(component_function, UNARY_FUNCTION_ARGUMENT_COUNT),
                    operator_node(AstNodeType::EqualToNode),
                ]);
            }
            nodes.push(function_node(
                LOGICAL_AND_FUNCTION_INDEX,
                TERNARY_FUNCTION_ARGUMENT_COUNT,
            ));
            nodes
        },
        FixedDatePredicateKind::Before => vec![
            linked_cell_node(formula_owner_uuid),
            date_node(lower),
            operator_node(AstNodeType::LessThanNode),
            linked_cell_node(formula_owner_uuid),
            function_node(IS_BLANK_FUNCTION_INDEX, UNARY_FUNCTION_ARGUMENT_COUNT),
            function_node(LOGICAL_NOT_FUNCTION_INDEX, UNARY_FUNCTION_ARGUMENT_COUNT),
            function_node(LOGICAL_AND_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
        ],
        FixedDatePredicateKind::After => vec![
            linked_cell_node(formula_owner_uuid),
            date_node(lower),
            duration_node(ONE_DAY),
            operator_node(AstNodeType::AdditionNode),
            operator_node(AstNodeType::GreaterThanOrEqualToNode),
        ],
        FixedDatePredicateKind::Between => {
            let upper = upper.expect("fixed date range requires an upper bound");
            vec![
                date_node(lower),
                date_node(upper),
                operator_node(AstNodeType::LessThanOrEqualToNode),
                operator_node(AstNodeType::BeginThunkNode),
                linked_cell_node(formula_owner_uuid),
                date_node(lower),
                operator_node(AstNodeType::GreaterThanOrEqualToNode),
                linked_cell_node(formula_owner_uuid),
                date_node(upper),
                duration_node(ONE_DAY),
                operator_node(AstNodeType::AdditionNode),
                operator_node(AstNodeType::LessThanNode),
                function_node(LOGICAL_AND_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
                operator_node(AstNodeType::EndThunkNode),
                operator_node(AstNodeType::BeginThunkNode),
                linked_cell_node(formula_owner_uuid),
                date_node(upper),
                operator_node(AstNodeType::GreaterThanOrEqualToNode),
                linked_cell_node(formula_owner_uuid),
                date_node(lower),
                duration_node(ONE_DAY),
                operator_node(AstNodeType::AdditionNode),
                operator_node(AstNodeType::LessThanNode),
                function_node(LOGICAL_AND_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
                operator_node(AstNodeType::EndThunkNode),
                function_node(
                    CONDITIONAL_FUNCTION_INDEX,
                    CONDITIONAL_FUNCTION_ARGUMENT_COUNT,
                ),
            ]
        },
    }
}

pub(super) fn validate_fixed(
    formula: &tsce::FormulaArchive,
    kind: FixedDatePredicateKind,
    lower: f64,
    upper: Option<f64>,
) -> Result<()> {
    if kind.is_range() != upper.is_some() {
        return Err(invalid_formula());
    }
    let expected = fixed_nodes(kind, lower, upper, &tsp::Uuid { upper: 0, lower: 0 });
    validate_nodes(formula, &expected)
}

fn validate_nodes(
    formula: &tsce::FormulaArchive,
    expected: &[tsce::ast_node_array_archive::AstNodeArchive],
) -> Result<()> {
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

fn date_node(value: f64) -> tsce::ast_node_array_archive::AstNodeArchive {
    use tsce::ast_node_array_archive::AstNodeType;

    tsce::ast_node_array_archive::AstNodeArchive {
        ast_node_type: AstNodeType::DateNode as i32,
        ast_date_node_date_num: Some(value),
        ast_date_node_suppress_date_format: Some(false),
        ast_date_node_suppress_time_format: Some(false),
        ast_date_node_date_time_format: Some(String::new()),
        ..Default::default()
    }
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
