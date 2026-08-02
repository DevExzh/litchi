//! Canonical operand-free cell-predicate formula graphs.

use super::*;

pub(super) fn nodes(
    kind: CellPredicateKind,
    formula_owner_uuid: &tsp::Uuid,
) -> Vec<tsce::ast_node_array_archive::AstNodeArchive> {
    let mut nodes = vec![
        linked_cell_node(formula_owner_uuid),
        function_node(IS_BLANK_FUNCTION_INDEX, UNARY_FUNCTION_ARGUMENT_COUNT),
    ];
    if kind == CellPredicateKind::IsNotBlank {
        nodes.push(function_node(
            LOGICAL_NOT_FUNCTION_INDEX,
            UNARY_FUNCTION_ARGUMENT_COUNT,
        ));
    }
    nodes
}

pub(super) fn validate(formula: &tsce::FormulaArchive, kind: CellPredicateKind) -> Result<()> {
    let expected_len = match kind {
        CellPredicateKind::IsBlank => 2,
        CellPredicateKind::IsNotBlank => 3,
    };
    let nodes = &formula.ast_node_array.ast_node;
    if nodes.len() != expected_len
        || !node_matches(
            &nodes[0],
            tsce::ast_node_array_archive::AstNodeType::LinkedCellRefNode,
        )
        || !function_matches(
            &nodes[1],
            IS_BLANK_FUNCTION_INDEX,
            UNARY_FUNCTION_ARGUMENT_COUNT,
        )
        || (kind == CellPredicateKind::IsNotBlank
            && !function_matches(
                &nodes[2],
                LOGICAL_NOT_FUNCTION_INDEX,
                UNARY_FUNCTION_ARGUMENT_COUNT,
            ))
    {
        return Err(invalid_formula());
    }
    Ok(())
}
