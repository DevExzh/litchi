//! Canonical operand-free Checkbox-predicate formula graphs.

use super::*;

pub(super) fn nodes(
    kind: CheckboxPredicateKind,
    formula_owner_uuid: &tsp::Uuid,
) -> Result<Vec<tsce::ast_node_array_archive::AstNodeArchive>> {
    use tsce::ast_node_array_archive::AstNodeType;

    let mut nodes = vec![linked_cell_node(formula_owner_uuid)];
    if !kind.is_checked() {
        nodes.push(function_node(
            LOGICAL_NOT_FUNCTION_INDEX,
            UNARY_FUNCTION_ARGUMENT_COUNT,
        ));
    }
    nodes.extend([
        linked_cell_node(formula_owner_uuid),
        function_node(
            CELL_DATA_FORMAT_FUNCTION_INDEX,
            UNARY_FUNCTION_ARGUMENT_COUNT,
        ),
        number_node(CHECKBOX_DATA_FORMAT_CODE)?,
        operator_node(AstNodeType::EqualToNode),
        function_node(LOGICAL_AND_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
        operator_node(AstNodeType::BeginThunkNode),
        boolean_node(true),
        operator_node(AstNodeType::EndThunkNode),
        operator_node(AstNodeType::BeginThunkNode),
        boolean_node(false),
        operator_node(AstNodeType::EndThunkNode),
        function_node(
            CONDITIONAL_FUNCTION_INDEX,
            CONDITIONAL_FUNCTION_ARGUMENT_COUNT,
        ),
    ]);
    Ok(nodes)
}

pub(super) fn prepivot_nodes(
    kind: CheckboxPredicateKind,
    formula_owner_uuid: &tsp::Uuid,
) -> Vec<tsce::ast_node_array_archive::AstNodeArchive> {
    vec![
        linked_cell_node(formula_owner_uuid),
        boolean_node(kind.is_checked()),
        operator_node(tsce::ast_node_array_archive::AstNodeType::EqualToNode),
    ]
}

pub(super) fn validate(formula: &tsce::FormulaArchive, kind: CheckboxPredicateKind) -> Result<()> {
    use tsce::ast_node_array_archive::AstNodeType;

    let nodes = formula.ast_node_array.ast_node.as_slice();
    let (value_cell, format_nodes) = if kind.is_checked() {
        let [
            value_cell,
            format_cell,
            data_format,
            checkbox_format,
            equal,
            logical_and,
            true_begin,
            true_value,
            true_end,
            false_begin,
            false_value,
            false_end,
            conditional,
        ] = nodes
        else {
            return Err(invalid_formula());
        };
        (
            value_cell,
            [
                format_cell,
                data_format,
                checkbox_format,
                equal,
                logical_and,
                true_begin,
                true_value,
                true_end,
                false_begin,
                false_value,
                false_end,
                conditional,
            ],
        )
    } else {
        let [
            value_cell,
            logical_not,
            format_cell,
            data_format,
            checkbox_format,
            equal,
            logical_and,
            true_begin,
            true_value,
            true_end,
            false_begin,
            false_value,
            false_end,
            conditional,
        ] = nodes
        else {
            return Err(invalid_formula());
        };
        if !function_matches(
            logical_not,
            LOGICAL_NOT_FUNCTION_INDEX,
            UNARY_FUNCTION_ARGUMENT_COUNT,
        ) {
            return Err(invalid_formula());
        }
        (
            value_cell,
            [
                format_cell,
                data_format,
                checkbox_format,
                equal,
                logical_and,
                true_begin,
                true_value,
                true_end,
                false_begin,
                false_value,
                false_end,
                conditional,
            ],
        )
    };
    let [
        format_cell,
        data_format,
        checkbox_format,
        equal,
        logical_and,
        true_begin,
        true_value,
        true_end,
        false_begin,
        false_value,
        false_end,
        conditional,
    ] = format_nodes;
    if !node_matches(value_cell, AstNodeType::LinkedCellRefNode)
        || !node_matches(format_cell, AstNodeType::LinkedCellRefNode)
        || !function_matches(
            data_format,
            CELL_DATA_FORMAT_FUNCTION_INDEX,
            UNARY_FUNCTION_ARGUMENT_COUNT,
        )
        || !number_matches(checkbox_format, CHECKBOX_DATA_FORMAT_CODE)
        || !node_matches(equal, AstNodeType::EqualToNode)
        || !function_matches(
            logical_and,
            LOGICAL_AND_FUNCTION_INDEX,
            BINARY_FUNCTION_ARGUMENT_COUNT,
        )
        || !node_matches(true_begin, AstNodeType::BeginThunkNode)
        || !boolean_matches(true_value, true)
        || !node_matches(true_end, AstNodeType::EndThunkNode)
        || !node_matches(false_begin, AstNodeType::BeginThunkNode)
        || !boolean_matches(false_value, false)
        || !node_matches(false_end, AstNodeType::EndThunkNode)
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

pub(super) fn validate_prepivot(
    formula: &tsce::FormulaArchive,
    kind: CheckboxPredicateKind,
) -> Result<()> {
    use tsce::ast_node_array_archive::AstNodeType;

    let [cell, value, equal] = formula.ast_node_array.ast_node.as_slice() else {
        return Err(invalid_formula());
    };
    if !node_matches(cell, AstNodeType::LinkedCellRefNode)
        || !boolean_matches(value, kind.is_checked())
        || !node_matches(equal, AstNodeType::EqualToNode)
    {
        return Err(invalid_formula());
    }
    Ok(())
}

fn boolean_matches(node: &tsce::ast_node_array_archive::AstNodeArchive, value: bool) -> bool {
    node_matches(node, tsce::ast_node_array_archive::AstNodeType::BooleanNode)
        && node.ast_boolean_node_boolean == Some(value)
}
