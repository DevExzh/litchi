//! Native formula graphs for date-period conditional-highlight predicates.

use super::*;
use crate::numbers::editor::conditional_highlight::native::{
    DATE_ADD_MONTHS_FUNCTION_INDEX, DATE_DIFFERENCE_FUNCTION_INDEX,
    DATE_DURATION_FROM_WEEKS_DAYS_FUNCTION_INDEX, DatePeriodPredicateKind,
};
use litchi_iwa_common::table::cell::conditional_highlight::{
    Offset, OffsetDirection, Period, PeriodUnit,
};

const ZERO: f64 = 0.0;
const ONE_DAY: u32 = 1;
const MONTHS_PER_QUARTER: f64 = 3.0;
const MONTHS_PER_YEAR: f64 = 12.0;
const FORWARD_SIGN: f64 = 1.0;
const BACKWARD_SIGN: f64 = -1.0;
const DATE_DIFFERENCE_UNIT: &str = "D";

const NEXT_DAYS_QUANTITY_INDEX: i32 = 6;
const NEXT_OTHER_QUANTITY_INDEX: i32 = 5;
const LAST_DAYS_QUANTITY_INDEX: i32 = 3;
const LAST_OTHER_QUANTITY_INDEX: i32 = 2;
const OFFSET_DAYS_QUANTITY_INDEX: i32 = 3;
const OFFSET_OTHER_QUANTITY_INDEX: i32 = 2;

pub(in crate::numbers::editor::conditional_highlight::formula) fn quantity_node_index(
    kind: DatePeriodPredicateKind,
    unit: PeriodUnit,
) -> i32 {
    match (kind, unit) {
        (DatePeriodPredicateKind::InNext, PeriodUnit::Days) => NEXT_DAYS_QUANTITY_INDEX,
        (DatePeriodPredicateKind::InNext, _) => NEXT_OTHER_QUANTITY_INDEX,
        (DatePeriodPredicateKind::InLast, PeriodUnit::Days) => LAST_DAYS_QUANTITY_INDEX,
        (DatePeriodPredicateKind::InLast, _) => LAST_OTHER_QUANTITY_INDEX,
        (DatePeriodPredicateKind::OffsetFromToday, PeriodUnit::Days) => OFFSET_DAYS_QUANTITY_INDEX,
        (DatePeriodPredicateKind::OffsetFromToday, _) => OFFSET_OTHER_QUANTITY_INDEX,
    }
}

pub(in crate::numbers::editor::conditional_highlight::formula) fn nodes(
    kind: DatePeriodPredicateKind,
    period: Period,
    direction: Option<OffsetDirection>,
    formula_owner_uuid: &tsp::Uuid,
) -> Result<Vec<tsce::ast_node_array_archive::AstNodeArchive>> {
    match kind {
        DatePeriodPredicateKind::InNext => range_nodes(period, true, formula_owner_uuid),
        DatePeriodPredicateKind::InLast => range_nodes(period, false, formula_owner_uuid),
        DatePeriodPredicateKind::OffsetFromToday => exact_nodes(
            Offset::new(period, direction.ok_or_else(invalid_formula)?),
            formula_owner_uuid,
        ),
    }
}

pub(in crate::numbers::editor::conditional_highlight::formula) fn validate(
    formula: &tsce::FormulaArchive,
    kind: DatePeriodPredicateKind,
    period: Period,
    direction: Option<OffsetDirection>,
) -> Result<()> {
    let expected = nodes(kind, period, direction, &tsp::Uuid { upper: 0, lower: 0 })?;
    validate_nodes(formula, &expected)
}

fn range_nodes(
    period: Period,
    forward: bool,
    formula_owner_uuid: &tsp::Uuid,
) -> Result<Vec<tsce::ast_node_array_archive::AstNodeArchive>> {
    use tsce::ast_node_array_archive::AstNodeType;

    let mut nodes = Vec::with_capacity(16);
    if forward {
        nodes.extend([
            linked_cell_node(formula_owner_uuid),
            today_node(),
            operator_node(AstNodeType::GreaterThanOrEqualToNode),
            linked_cell_node(formula_owner_uuid),
        ]);
        nodes.extend(shifted_today_nodes(period, FORWARD_SIGN)?);
        nodes.extend([
            duration_node(ONE_DAY),
            operator_node(AstNodeType::AdditionNode),
            operator_node(AstNodeType::LessThanNode),
        ]);
    } else {
        nodes.push(linked_cell_node(formula_owner_uuid));
        nodes.extend(shifted_today_nodes(period, BACKWARD_SIGN)?);
        nodes.extend([
            operator_node(AstNodeType::GreaterThanOrEqualToNode),
            linked_cell_node(formula_owner_uuid),
            today_node(),
            duration_node(ONE_DAY),
            operator_node(AstNodeType::AdditionNode),
            operator_node(AstNodeType::LessThanNode),
        ]);
    }
    nodes.push(function_node(
        LOGICAL_AND_FUNCTION_INDEX,
        BINARY_FUNCTION_ARGUMENT_COUNT,
    ));
    Ok(nodes)
}

fn exact_nodes(
    offset: Offset,
    formula_owner_uuid: &tsp::Uuid,
) -> Result<Vec<tsce::ast_node_array_archive::AstNodeArchive>> {
    use tsce::ast_node_array_archive::AstNodeType;

    let sign = match offset.direction() {
        OffsetDirection::Ago => BACKWARD_SIGN,
        OffsetDirection::FromNow => FORWARD_SIGN,
    };
    let mut nodes = Vec::with_capacity(14);
    nodes.push(linked_cell_node(formula_owner_uuid));
    nodes.extend(shifted_today_nodes(offset.period(), sign)?);
    nodes.extend([
        string_node(DATE_DIFFERENCE_UNIT),
        function_node(
            DATE_DIFFERENCE_FUNCTION_INDEX,
            TERNARY_FUNCTION_ARGUMENT_COUNT,
        ),
        number_node(ZERO)?,
        operator_node(AstNodeType::EqualToNode),
    ]);
    Ok(nodes)
}

fn shifted_today_nodes(
    period: Period,
    sign: f64,
) -> Result<Vec<tsce::ast_node_array_archive::AstNodeArchive>> {
    use PeriodUnit as Unit;
    use tsce::ast_node_array_archive::AstNodeType;

    let count = f64::from(period.count());
    let mut nodes = vec![today_node()];
    match period.unit() {
        Unit::Days => nodes.extend([
            number_node(ZERO)?,
            number_node(count)?,
            function_node(
                DATE_DURATION_FROM_WEEKS_DAYS_FUNCTION_INDEX,
                BINARY_FUNCTION_ARGUMENT_COUNT,
            ),
            number_node(sign)?,
            operator_node(AstNodeType::MultiplicationNode),
            operator_node(AstNodeType::AdditionNode),
        ]),
        Unit::Weeks => nodes.extend([
            number_node(count)?,
            number_node(ZERO)?,
            function_node(
                DATE_DURATION_FROM_WEEKS_DAYS_FUNCTION_INDEX,
                BINARY_FUNCTION_ARGUMENT_COUNT,
            ),
            number_node(sign)?,
            operator_node(AstNodeType::MultiplicationNode),
            operator_node(AstNodeType::AdditionNode),
        ]),
        Unit::Months | Unit::Quarters | Unit::Years => {
            nodes.push(number_node(count)?);
            if let Some(multiplier) = month_multiplier(period.unit()) {
                nodes.extend([
                    number_node(multiplier)?,
                    operator_node(AstNodeType::MultiplicationNode),
                ]);
            }
            nodes.extend([
                number_node(sign)?,
                operator_node(AstNodeType::MultiplicationNode),
                function_node(
                    DATE_ADD_MONTHS_FUNCTION_INDEX,
                    BINARY_FUNCTION_ARGUMENT_COUNT,
                ),
            ]);
        },
    }
    Ok(nodes)
}

fn month_multiplier(unit: PeriodUnit) -> Option<f64> {
    match unit {
        PeriodUnit::Quarters => Some(MONTHS_PER_QUARTER),
        PeriodUnit::Years => Some(MONTHS_PER_YEAR),
        _ => None,
    }
}

fn string_node(value: &str) -> tsce::ast_node_array_archive::AstNodeArchive {
    use tsce::ast_node_array_archive::AstNodeType;

    tsce::ast_node_array_archive::AstNodeArchive {
        ast_node_type: AstNodeType::StringNode as i32,
        ast_string_node_string: Some(value.to_owned()),
        ..Default::default()
    }
}
