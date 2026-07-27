//! Canonical formula graphs used by conditional-highlight predicates.

mod boolean;
mod cell;
mod sign;
mod text;

use super::*;

pub(super) fn encode(
    condition: &TableCellConditionalHighlightCondition,
    formula_owner_uuid: &tsp::Uuid,
) -> Result<tsce::FormulaArchive> {
    let kind = NativePredicateKind::from_condition(condition);
    let nodes = match (kind, condition) {
        (NativePredicateKind::Cell(kind), _) => cell::nodes(kind, formula_owner_uuid),
        (NativePredicateKind::Boolean(kind), _) => boolean::nodes(kind, formula_owner_uuid)?,
        (NativePredicateKind::NumericSign(kind), _) => sign::nodes(kind, formula_owner_uuid)?,
        (
            NativePredicateKind::Numeric(kind),
            TableCellConditionalHighlightCondition::Between(range)
            | TableCellConditionalHighlightCondition::NotBetween(range),
        ) => range_nodes(
            kind,
            range.lower().get(),
            range.upper().get(),
            formula_owner_uuid,
        )?,
        (NativePredicateKind::Numeric(kind), _) => {
            let value = condition
                .single_operand()
                .expect("single-number predicate has one operand")
                .get();
            vec![
                linked_cell_node(formula_owner_uuid),
                number_node(value)?,
                operator_node(
                    kind.single_ast_node_type()
                        .expect("single-number predicate has a comparison node"),
                ),
            ]
        },
        (NativePredicateKind::Text(kind), _) => text::nodes(
            kind,
            condition
                .text()
                .expect("text predicate has a text operand")
                .as_str(),
            formula_owner_uuid,
        )?,
    };
    Ok(tsce::FormulaArchive {
        ast_node_array: tsce::AstNodeArrayArchive { ast_node: nodes },
        ..Default::default()
    })
}

pub(super) fn encode_prepivot(
    condition: &TableCellConditionalHighlightCondition,
    formula_owner_uuid: &tsp::Uuid,
) -> Result<tsce::FormulaArchive> {
    let kind = NativePredicateKind::from_condition(condition);
    if let NativePredicateKind::Boolean(kind) = kind {
        return Ok(tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive {
                ast_node: boolean::prepivot_nodes(kind, formula_owner_uuid),
            },
            ..Default::default()
        });
    }
    if let NativePredicateKind::NumericSign(kind) = kind {
        return Ok(tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive {
                ast_node: sign::prepivot_nodes(kind, formula_owner_uuid)?,
            },
            ..Default::default()
        });
    }
    encode(condition, formula_owner_uuid)
}

pub(super) fn validate(
    formula: &tsce::FormulaArchive,
    kind: NativePredicateKind,
    condition: &TableCellConditionalHighlightCondition,
) -> Result<()> {
    match (kind, condition) {
        (NativePredicateKind::Cell(kind), _) => cell::validate(formula, kind),
        (NativePredicateKind::Boolean(kind), _) => boolean::validate(formula, kind),
        (NativePredicateKind::NumericSign(kind), _) => sign::validate(formula, kind),
        (
            NativePredicateKind::Numeric(kind),
            TableCellConditionalHighlightCondition::Between(range)
            | TableCellConditionalHighlightCondition::NotBetween(range),
        ) => validate_range(formula, kind, range.lower().get(), range.upper().get()),
        (NativePredicateKind::Numeric(kind), _) => {
            let value = condition
                .single_operand()
                .ok_or_else(invalid_formula)?
                .get();
            validate_single_number(formula, kind, value)
        },
        (NativePredicateKind::Text(kind), _) => text::validate(
            formula,
            kind,
            condition.text().ok_or_else(invalid_formula)?.as_str(),
        ),
    }
}

pub(super) fn validate_prepivot(
    formula: &tsce::FormulaArchive,
    kind: NativePredicateKind,
    condition: &TableCellConditionalHighlightCondition,
) -> Result<()> {
    match kind {
        NativePredicateKind::Boolean(kind) => return boolean::validate_prepivot(formula, kind),
        NativePredicateKind::NumericSign(kind) => return sign::validate_prepivot(formula, kind),
        _ => {},
    }
    validate(formula, kind, condition)
}

fn linked_cell_node(
    formula_owner_uuid: &tsp::Uuid,
) -> tsce::ast_node_array_archive::AstNodeArchive {
    use tsce::ast_node_array_archive::AstNodeType;

    tsce::ast_node_array_archive::AstNodeArchive {
        ast_node_type: AstNodeType::LinkedCellRefNode as i32,
        ast_cross_table_reference_extra_info: Some(
            tsce::ast_node_array_archive::AstCrossTableReferenceExtraInfoArchive {
                table_id: uuid_as_cfuuid(formula_owner_uuid),
                ..Default::default()
            },
        ),
        ..Default::default()
    }
}

fn number_node(value: f64) -> Result<tsce::ast_node_array_archive::AstNodeArchive> {
    use tsce::ast_node_array_archive::AstNodeType;

    let decimal = crate::numbers::bnc::decimal128_le(value)?;
    Ok(tsce::ast_node_array_archive::AstNodeArchive {
        ast_node_type: AstNodeType::NumberNode as i32,
        ast_number_node_number: Some(value),
        ast_number_node_decimal_low: Some(u64::from_le_bytes(
            decimal[..8]
                .try_into()
                .expect("fixed-size decimal lower half"),
        )),
        ast_number_node_decimal_high: Some(u64::from_le_bytes(
            decimal[8..]
                .try_into()
                .expect("fixed-size decimal upper half"),
        )),
        ..Default::default()
    })
}

fn boolean_node(value: bool) -> tsce::ast_node_array_archive::AstNodeArchive {
    use tsce::ast_node_array_archive::AstNodeType;

    tsce::ast_node_array_archive::AstNodeArchive {
        ast_node_type: AstNodeType::BooleanNode as i32,
        ast_boolean_node_boolean: Some(value),
        ..Default::default()
    }
}

fn operator_node(
    node_type: tsce::ast_node_array_archive::AstNodeType,
) -> tsce::ast_node_array_archive::AstNodeArchive {
    tsce::ast_node_array_archive::AstNodeArchive {
        ast_node_type: node_type as i32,
        ..Default::default()
    }
}

fn function_node(index: u32, argument_count: u32) -> tsce::ast_node_array_archive::AstNodeArchive {
    use tsce::ast_node_array_archive::AstNodeType;

    tsce::ast_node_array_archive::AstNodeArchive {
        ast_node_type: AstNodeType::FunctionNode as i32,
        ast_function_node_index: Some(index),
        ast_function_node_num_args: Some(argument_count),
        ..Default::default()
    }
}

fn range_nodes(
    kind: NumericPredicateKind,
    lower: f64,
    upper: f64,
    formula_owner_uuid: &tsp::Uuid,
) -> Result<Vec<tsce::ast_node_array_archive::AstNodeArchive>> {
    use tsce::ast_node_array_archive::AstNodeType;

    let (
        first_lower_comparison,
        first_upper_comparison,
        second_lower_comparison,
        second_upper_comparison,
        logical_function,
    ) = match kind {
        NumericPredicateKind::Between => (
            AstNodeType::GreaterThanOrEqualToNode,
            AstNodeType::LessThanOrEqualToNode,
            AstNodeType::GreaterThanOrEqualToNode,
            AstNodeType::LessThanOrEqualToNode,
            LOGICAL_AND_FUNCTION_INDEX,
        ),
        NumericPredicateKind::NotBetween => (
            AstNodeType::LessThanNode,
            AstNodeType::GreaterThanNode,
            AstNodeType::LessThanNode,
            AstNodeType::GreaterThanNode,
            LOGICAL_OR_FUNCTION_INDEX,
        ),
        _ => unreachable!("range formula requires a range predicate"),
    };
    Ok(vec![
        number_node(lower)?,
        number_node(upper)?,
        operator_node(AstNodeType::LessThanOrEqualToNode),
        operator_node(AstNodeType::BeginThunkNode),
        linked_cell_node(formula_owner_uuid),
        number_node(lower)?,
        operator_node(first_lower_comparison),
        linked_cell_node(formula_owner_uuid),
        number_node(upper)?,
        operator_node(first_upper_comparison),
        function_node(logical_function, BINARY_FUNCTION_ARGUMENT_COUNT),
        operator_node(AstNodeType::EndThunkNode),
        operator_node(AstNodeType::BeginThunkNode),
        linked_cell_node(formula_owner_uuid),
        number_node(upper)?,
        operator_node(second_lower_comparison),
        linked_cell_node(formula_owner_uuid),
        number_node(lower)?,
        operator_node(second_upper_comparison),
        function_node(logical_function, BINARY_FUNCTION_ARGUMENT_COUNT),
        operator_node(AstNodeType::EndThunkNode),
        function_node(
            CONDITIONAL_FUNCTION_INDEX,
            CONDITIONAL_FUNCTION_ARGUMENT_COUNT,
        ),
    ])
}

fn validate_single_number(
    formula: &tsce::FormulaArchive,
    kind: NumericPredicateKind,
    value: f64,
) -> Result<()> {
    use tsce::ast_node_array_archive::AstNodeType;

    let comparison = kind.single_ast_node_type().ok_or_else(invalid_formula)?;
    let [cell, number, operator] = formula.ast_node_array.ast_node.as_slice() else {
        return Err(invalid_formula());
    };
    if !node_matches(cell, AstNodeType::LinkedCellRefNode)
        || !number_matches(number, value)
        || !node_matches(operator, comparison)
    {
        return Err(invalid_formula());
    }
    Ok(())
}

fn validate_range(
    formula: &tsce::FormulaArchive,
    kind: NumericPredicateKind,
    lower: f64,
    upper: f64,
) -> Result<()> {
    use tsce::ast_node_array_archive::AstNodeType;

    let (lower_comparison, upper_comparison, logical_function) = match kind {
        NumericPredicateKind::Between => (
            AstNodeType::GreaterThanOrEqualToNode,
            AstNodeType::LessThanOrEqualToNode,
            LOGICAL_AND_FUNCTION_INDEX,
        ),
        NumericPredicateKind::NotBetween => (
            AstNodeType::LessThanNode,
            AstNodeType::GreaterThanNode,
            LOGICAL_OR_FUNCTION_INDEX,
        ),
        _ => return Err(invalid_formula()),
    };
    let [
        ordered_lower,
        ordered_upper,
        order_operator,
        first_begin,
        first_lower_cell,
        first_lower_bound,
        first_lower_operator,
        first_upper_cell,
        first_upper_bound,
        first_upper_operator,
        first_logical,
        first_end,
        second_begin,
        second_lower_cell,
        second_lower_bound,
        second_lower_operator,
        second_upper_cell,
        second_upper_bound,
        second_upper_operator,
        second_logical,
        second_end,
        conditional,
    ] = formula.ast_node_array.ast_node.as_slice()
    else {
        return Err(invalid_formula());
    };
    if !number_matches(ordered_lower, lower)
        || !number_matches(ordered_upper, upper)
        || !node_matches(order_operator, AstNodeType::LessThanOrEqualToNode)
        || !node_matches(first_begin, AstNodeType::BeginThunkNode)
        || !node_matches(first_lower_cell, AstNodeType::LinkedCellRefNode)
        || !number_matches(first_lower_bound, lower)
        || !node_matches(first_lower_operator, lower_comparison)
        || !node_matches(first_upper_cell, AstNodeType::LinkedCellRefNode)
        || !number_matches(first_upper_bound, upper)
        || !node_matches(first_upper_operator, upper_comparison)
        || !function_matches(
            first_logical,
            logical_function,
            BINARY_FUNCTION_ARGUMENT_COUNT,
        )
        || !node_matches(first_end, AstNodeType::EndThunkNode)
        || !node_matches(second_begin, AstNodeType::BeginThunkNode)
        || !node_matches(second_lower_cell, AstNodeType::LinkedCellRefNode)
        || !number_matches(second_lower_bound, upper)
        || !node_matches(second_lower_operator, lower_comparison)
        || !node_matches(second_upper_cell, AstNodeType::LinkedCellRefNode)
        || !number_matches(second_upper_bound, lower)
        || !node_matches(second_upper_operator, upper_comparison)
        || !function_matches(
            second_logical,
            logical_function,
            BINARY_FUNCTION_ARGUMENT_COUNT,
        )
        || !node_matches(second_end, AstNodeType::EndThunkNode)
        || !function_matches(
            conditional,
            CONDITIONAL_FUNCTION_INDEX,
            CONDITIONAL_FUNCTION_ARGUMENT_COUNT,
        )
    {
        return Err(invalid_formula());
    }
    Ok(())
}

fn node_matches(
    node: &tsce::ast_node_array_archive::AstNodeArchive,
    node_type: tsce::ast_node_array_archive::AstNodeType,
) -> bool {
    node.ast_node_type == node_type as i32
}

fn number_matches(node: &tsce::ast_node_array_archive::AstNodeArchive, value: f64) -> bool {
    use tsce::ast_node_array_archive::AstNodeType;

    node_matches(node, AstNodeType::NumberNode) && node.ast_number_node_number == Some(value)
}

fn function_matches(
    node: &tsce::ast_node_array_archive::AstNodeArchive,
    index: u32,
    argument_count: u32,
) -> bool {
    use tsce::ast_node_array_archive::AstNodeType;

    node_matches(node, AstNodeType::FunctionNode)
        && node.ast_function_node_index == Some(index)
        && node.ast_function_node_num_args == Some(argument_count)
}

fn invalid_formula() -> Error {
    Error::InvalidFormat(
        "iWork conditional-highlight rule uses an unsupported native graph".to_owned(),
    )
}
