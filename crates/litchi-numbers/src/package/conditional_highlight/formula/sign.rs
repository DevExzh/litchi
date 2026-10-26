//! Canonical operand-free numeric-sign predicate formula graphs.

use super::*;

pub(super) fn nodes(
    kind: NumericSignPredicateKind,
    formula_owner_uuid: &tsp::Uuid,
) -> Result<Vec<tsce::ast_node_array_archive::AstNodeArchive>> {
    Ok(vec![
        linked_cell_node(formula_owner_uuid),
        number_node(0.0)?,
        operator_node(kind.comparison()),
        linked_cell_node(formula_owner_uuid),
        function_node(IS_NUMBER_FUNCTION_INDEX, UNARY_FUNCTION_ARGUMENT_COUNT),
        function_node(LOGICAL_AND_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
    ])
}

pub(super) fn prepivot_nodes(
    kind: NumericSignPredicateKind,
    formula_owner_uuid: &tsp::Uuid,
) -> Result<Vec<tsce::ast_node_array_archive::AstNodeArchive>> {
    Ok(vec![
        linked_cell_node(formula_owner_uuid),
        number_node(0.0)?,
        operator_node(kind.comparison()),
    ])
}

pub(super) fn validate(
    formula: &tsce::FormulaArchive,
    kind: NumericSignPredicateKind,
) -> Result<()> {
    let [
        first_cell,
        zero,
        comparison,
        second_cell,
        is_number,
        logical_and,
    ] = formula.ast_node_array.ast_node.as_slice()
    else {
        return Err(invalid_formula());
    };
    if !node_matches(
        first_cell,
        tsce::ast_node_array_archive::AstNodeType::LinkedCellRefNode,
    ) || !number_matches(zero, 0.0)
        || !node_matches(comparison, kind.comparison())
        || !node_matches(
            second_cell,
            tsce::ast_node_array_archive::AstNodeType::LinkedCellRefNode,
        )
        || !function_matches(
            is_number,
            IS_NUMBER_FUNCTION_INDEX,
            UNARY_FUNCTION_ARGUMENT_COUNT,
        )
        || !function_matches(
            logical_and,
            LOGICAL_AND_FUNCTION_INDEX,
            BINARY_FUNCTION_ARGUMENT_COUNT,
        )
    {
        return Err(invalid_formula());
    }
    Ok(())
}

pub(super) fn validate_prepivot(
    formula: &tsce::FormulaArchive,
    kind: NumericSignPredicateKind,
) -> Result<()> {
    let [cell, zero, comparison] = formula.ast_node_array.ast_node.as_slice() else {
        return Err(invalid_formula());
    };
    if !node_matches(
        cell,
        tsce::ast_node_array_archive::AstNodeType::LinkedCellRefNode,
    ) || !number_matches(zero, 0.0)
        || !node_matches(comparison, kind.comparison())
    {
        return Err(invalid_formula());
    }
    Ok(())
}
