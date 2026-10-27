//! Canonical text-predicate formula graphs.

use super::*;

#[derive(Clone, Copy)]
enum FormulaToken<'a> {
    String(&'a str),
    LinkedCell,
    Number(f64),
    Boolean(bool),
    Operator(tsce::ast_node_array_archive::AstNodeType),
    Function(u32, u32),
}

pub(super) fn nodes(
    kind: TextPredicateKind,
    text: &str,
    formula_owner_uuid: &tsp::Uuid,
) -> Result<Vec<tsce::ast_node_array_archive::AstNodeArchive>> {
    formula_tokens(kind, text)
        .into_iter()
        .map(|token| match token {
            FormulaToken::String(value) => Ok(string_node(value)),
            FormulaToken::LinkedCell => Ok(linked_cell_node(formula_owner_uuid)),
            FormulaToken::Number(value) => number_node(value),
            FormulaToken::Boolean(value) => Ok(boolean_node(value)),
            FormulaToken::Operator(kind) => Ok(operator_node(kind)),
            FormulaToken::Function(index, arguments) => Ok(function_node(index, arguments)),
        })
        .collect()
}

pub(super) fn validate(
    formula: &tsce::FormulaArchive,
    kind: TextPredicateKind,
    text: &str,
) -> Result<()> {
    let expected = formula_tokens(kind, text);
    if formula.ast_node_array.ast_node.len() != expected.len()
        || !formula
            .ast_node_array
            .ast_node
            .iter()
            .zip(expected)
            .all(|(node, token)| token_matches(node, token))
    {
        return Err(invalid_formula());
    }
    Ok(())
}

fn formula_tokens(kind: TextPredicateKind, text: &str) -> Vec<FormulaToken<'_>> {
    use FormulaToken::{Boolean, Function, LinkedCell, Number, Operator, String};
    use tsce::ast_node_array_archive::AstNodeType;

    match kind {
        TextPredicateKind::EqualTo => vec![
            String(text),
            Function(TEXT_LENGTH_FUNCTION_INDEX, UNARY_FUNCTION_ARGUMENT_COUNT),
            LinkedCell,
            Function(TEXT_LENGTH_FUNCTION_INDEX, UNARY_FUNCTION_ARGUMENT_COUNT),
            Operator(AstNodeType::EqualToNode),
            Operator(AstNodeType::BeginThunkNode),
            String(text),
            LinkedCell,
            Function(TEXT_SEARCH_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
            Number(1.0),
            Operator(AstNodeType::EqualToNode),
            Operator(AstNodeType::EndThunkNode),
            Function(CONDITIONAL_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
        ],
        TextPredicateKind::NotEqualTo => vec![
            String(text),
            Function(TEXT_LENGTH_FUNCTION_INDEX, UNARY_FUNCTION_ARGUMENT_COUNT),
            LinkedCell,
            Function(TEXT_LENGTH_FUNCTION_INDEX, UNARY_FUNCTION_ARGUMENT_COUNT),
            Operator(AstNodeType::NotEqualToNode),
            Operator(AstNodeType::BeginThunkNode),
            Boolean(true),
            Operator(AstNodeType::EndThunkNode),
            Operator(AstNodeType::BeginThunkNode),
            String(text),
            LinkedCell,
            Function(TEXT_SEARCH_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
            Function(IS_ERROR_FUNCTION_INDEX, UNARY_FUNCTION_ARGUMENT_COUNT),
            Operator(AstNodeType::BeginThunkNode),
            Boolean(true),
            Operator(AstNodeType::EndThunkNode),
            Operator(AstNodeType::BeginThunkNode),
            String(text),
            LinkedCell,
            Function(TEXT_SEARCH_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
            Number(1.0),
            Operator(AstNodeType::NotEqualToNode),
            Operator(AstNodeType::EndThunkNode),
            Function(
                CONDITIONAL_FUNCTION_INDEX,
                CONDITIONAL_FUNCTION_ARGUMENT_COUNT,
            ),
            Operator(AstNodeType::EndThunkNode),
            Function(
                CONDITIONAL_FUNCTION_INDEX,
                CONDITIONAL_FUNCTION_ARGUMENT_COUNT,
            ),
        ],
        TextPredicateKind::StartsWith => vec![
            String(text),
            LinkedCell,
            Function(TEXT_SEARCH_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
            Number(1.0),
            Operator(AstNodeType::EqualToNode),
        ],
        TextPredicateKind::DoesNotStartWith => vec![
            String(text),
            LinkedCell,
            Function(TEXT_SEARCH_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
            Number(1.0),
            Operator(AstNodeType::NotEqualToNode),
            Boolean(true),
            Function(IF_ERROR_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
        ],
        TextPredicateKind::EndsWith => vec![
            String(text),
            LinkedCell,
            String(text),
            Function(TEXT_LENGTH_FUNCTION_INDEX, UNARY_FUNCTION_ARGUMENT_COUNT),
            Function(TEXT_RIGHT_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
            Function(TEXT_SEARCH_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
        ],
        TextPredicateKind::DoesNotEndWith => vec![
            String(text),
            LinkedCell,
            String(text),
            Function(TEXT_LENGTH_FUNCTION_INDEX, UNARY_FUNCTION_ARGUMENT_COUNT),
            Function(TEXT_RIGHT_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
            Function(TEXT_SEARCH_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
            Function(LOGICAL_NOT_FUNCTION_INDEX, UNARY_FUNCTION_ARGUMENT_COUNT),
            Boolean(true),
            Function(IF_ERROR_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
        ],
        TextPredicateKind::Contains => vec![
            String(text),
            LinkedCell,
            Function(TEXT_SEARCH_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
            Function(IS_ERROR_FUNCTION_INDEX, UNARY_FUNCTION_ARGUMENT_COUNT),
            Function(LOGICAL_NOT_FUNCTION_INDEX, UNARY_FUNCTION_ARGUMENT_COUNT),
        ],
        TextPredicateKind::DoesNotContain => vec![
            String(text),
            LinkedCell,
            Function(TEXT_SEARCH_FUNCTION_INDEX, BINARY_FUNCTION_ARGUMENT_COUNT),
            Function(IS_ERROR_FUNCTION_INDEX, UNARY_FUNCTION_ARGUMENT_COUNT),
        ],
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

fn boolean_node(value: bool) -> tsce::ast_node_array_archive::AstNodeArchive {
    use tsce::ast_node_array_archive::AstNodeType;

    tsce::ast_node_array_archive::AstNodeArchive {
        ast_node_type: AstNodeType::BooleanNode as i32,
        ast_boolean_node_boolean: Some(value),
        ..Default::default()
    }
}

fn token_matches(
    node: &tsce::ast_node_array_archive::AstNodeArchive,
    token: FormulaToken<'_>,
) -> bool {
    use tsce::ast_node_array_archive::AstNodeType;

    match token {
        FormulaToken::String(value) => {
            node_matches(node, AstNodeType::StringNode)
                && node.ast_string_node_string.as_deref() == Some(value)
        },
        FormulaToken::LinkedCell => node_matches(node, AstNodeType::LinkedCellRefNode),
        FormulaToken::Number(value) => number_matches(node, value),
        FormulaToken::Boolean(value) => {
            node_matches(node, AstNodeType::BooleanNode)
                && node.ast_boolean_node_boolean == Some(value)
        },
        FormulaToken::Operator(kind) => node_matches(node, kind),
        FormulaToken::Function(index, arguments) => function_matches(node, index, arguments),
    }
}
