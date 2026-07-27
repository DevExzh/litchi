//! Canonical native object-graph encoding for conditional-highlight rules.

use prost::Message;

use super::*;

const CONDITIONAL_STYLE_SET_MESSAGE_TYPE: u32 = 6_010;
const CELL_STYLE_MESSAGE_TYPE: u32 = 6_004;
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const NATIVE_MESSAGE_VERSION: &[u32] = &[1, 0, 5];

pub(super) fn insert_conditional_style_graph(
    package: &mut IWorkPackage,
    archive_name: &str,
    style_set_id: u64,
    rules: &[TableCellConditionalHighlightRule],
    formula_owner_uuid: &tsp::Uuid,
) -> Result<()> {
    let mut prepivot = Vec::with_capacity(rules.len());
    let mut current = Vec::with_capacity(rules.len());
    let mut objects = Vec::with_capacity(1 + rules.len() * 2);
    let mut references = Vec::with_capacity(rules.len() * 2);
    for (index, rule) in rules.iter().copied().enumerate() {
        let offset = u64::try_from(index)
            .map_err(|_| Error::ParseError("conditional-highlight index exceeds u64".to_owned()))?
            .checked_mul(2)
            .ok_or_else(|| {
                Error::ParseError("conditional-highlight identifier overflow".to_owned())
            })?;
        let text_style_id = style_set_id
            .checked_add(offset)
            .and_then(|identifier| identifier.checked_add(1))
            .ok_or_else(|| {
                Error::ParseError("conditional-highlight identifier overflow".to_owned())
            })?;
        let cell_style_id = text_style_id.checked_add(1).ok_or_else(|| {
            Error::ParseError("conditional-highlight identifier overflow".to_owned())
        })?;
        let kind = NumericPredicateKind::from_condition(rule.condition);
        let formula = predicate_formula(rule.condition, formula_owner_uuid)?;
        let predicate_type = kind.native_value();
        let (first_value, second_value) = predicate_values(rule.condition);
        let (cell_index, first_index, second_index) = if kind.is_range() {
            (
                PREDICATE_RANGE_CELL_ARGUMENT_INDEX,
                PREDICATE_RANGE_LOWER_ARGUMENT_INDEX,
                PREDICATE_RANGE_UPPER_ARGUMENT_INDEX,
            )
        } else {
            (
                PREDICATE_CELL_ARGUMENT_INDEX,
                PREDICATE_NUMBER_ARGUMENT_INDEX,
                PREDICATE_UNUSED_ARGUMENT_INDEX,
            )
        };
        let cell_style = tsp::Reference {
            identifier: cell_style_id,
            ..Default::default()
        };
        let text_style = tsp::Reference {
            identifier: text_style_id,
            ..Default::default()
        };
        prepivot.push(
            tst::conditional_style_set_archive::ConditionalStyleRulePrePivot {
                predicate: tst::FormulaPredicatePrePivotArchive {
                    formula: formula.clone(),
                    predicate_type,
                    qualifier1: PREDICATE_QUALIFIER_NONE,
                    qualifier2: PREDICATE_QUALIFIER_NONE,
                    param_index1: first_index,
                    param_index2: second_index,
                    param_index0: cell_index,
                },
                cell_style,
                text_style,
            },
        );
        current.push(tst::conditional_style_set_archive::ConditionalStyleRule {
            predicate: Some(tst::FormulaPredicateArchive {
                predicate_type,
                qualifier1: PREDICATE_QUALIFIER_NONE,
                qualifier2: PREDICATE_QUALIFIER_NONE,
                param_value0: Some(tst::FormulaPredArgArchive {
                    arg_type: PREDICATE_ARGUMENT_RELATIVE_CELL,
                    relative_cell_ref: Some(tsce::RelativeCellRefArchive {
                        relative_row_offset: Some(0),
                        relative_column_offset: Some(0),
                        table_uid: Some(*formula_owner_uuid),
                        preserve_column: Some(false),
                        preserve_row: Some(false),
                        ..Default::default()
                    }),
                    ..Default::default()
                }),
                param_value1: Some(tst::FormulaPredArgArchive {
                    arg_type: PREDICATE_ARGUMENT_NUMBER,
                    arg_value: Some(predicate_number(first_value)?),
                    ..Default::default()
                }),
                param_value2: Some(match second_value {
                    Some(value) => tst::FormulaPredArgArchive {
                        arg_type: PREDICATE_ARGUMENT_NUMBER,
                        arg_value: Some(predicate_number(value)?),
                        ..Default::default()
                    },
                    None => tst::FormulaPredArgArchive {
                        arg_type: PREDICATE_ARGUMENT_NONE,
                        ..Default::default()
                    },
                }),
                formula: Some(formula),
                for_conditional_style: Some(true),
                ..Default::default()
            }),
            cell_style,
            text_style,
        });
        objects.push(paragraph_style_object(text_style_id, rule)?);
        objects.push(cell_style_object(cell_style_id, rule)?);
        references.push(cell_style_id);
        references.push(text_style_id);
    }
    let style_set = tst::ConditionalStyleSetArchive {
        rule_count: u32::try_from(rules.len()).map_err(|_| {
            Error::ParseError("conditional-highlight rule count exceeds u32".to_owned())
        })?,
        rules_prepivot: prepivot,
        rules: Some(tst::conditional_style_set_archive::ConditionalStyleRules { rule: current }),
    };
    let mut set_object = ArchiveObject::new(
        style_set_id,
        vec![RawMessage {
            type_: CONDITIONAL_STYLE_SET_MESSAGE_TYPE,
            data: style_set.encode_to_vec(),
        }],
    )?;
    set_object.archive_info.message_infos[0].versions = NATIVE_MESSAGE_VERSION.to_vec();
    set_object.archive_info.message_infos[0].object_references = references;
    package.update_archive(archive_name, |archive| {
        archive.insert_object(set_object)?;
        for object in objects {
            archive.insert_object(object)?;
        }
        Ok(())
    })?;
    Ok(())
}

fn paragraph_style_object(
    identifier: u64,
    rule: TableCellConditionalHighlightRule,
) -> Result<ArchiveObject> {
    let style = rule.style;
    let text_override_count = u32::from(style.bold()) + u32::from(style.text_color().is_some());
    let data = tswp::ParagraphStyleArchive {
        super_: tss::StyleArchive::default(),
        override_count: (text_override_count != 0).then_some(text_override_count),
        char_properties: (text_override_count != 0).then(|| {
            tswp::CharacterStylePropertiesArchive {
                bold: style.bold().then_some(true),
                font_color: style.text_color().map(crate::shapes::color_to_native),
                ..Default::default()
            }
        }),
        ..Default::default()
    }
    .encode_to_vec();
    style_object(identifier, PARAGRAPH_STYLE_MESSAGE_TYPE, data)
}

fn cell_style_object(
    identifier: u64,
    rule: TableCellConditionalHighlightRule,
) -> Result<ArchiveObject> {
    let fill = rule.style.fill().map(|color| tsd::FillArchive {
        color: Some(crate::shapes::color_to_native(color)),
        ..Default::default()
    });
    let data = tst::CellStyleArchive {
        super_: tss::StyleArchive::default(),
        override_count: fill.is_some().then_some(1),
        cell_properties: fill.map(|cell_fill| tst::CellStylePropertiesArchive {
            cell_fill: Some(cell_fill),
            ..Default::default()
        }),
    }
    .encode_to_vec();
    style_object(identifier, CELL_STYLE_MESSAGE_TYPE, data)
}

fn style_object(identifier: u64, message_type: u32, data: Vec<u8>) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data,
        }],
    )?;
    object.archive_info.message_infos[0].versions = NATIVE_MESSAGE_VERSION.to_vec();
    Ok(object)
}

fn predicate_formula(
    condition: TableCellConditionalHighlightCondition,
    formula_owner_uuid: &tsp::Uuid,
) -> Result<tsce::FormulaArchive> {
    let kind = NumericPredicateKind::from_condition(condition);
    let (first, second) = predicate_values(condition);
    let nodes = if let Some(second) = second {
        range_formula_nodes(kind, first, second, formula_owner_uuid)?
    } else {
        vec![
            linked_cell_node(formula_owner_uuid),
            number_node(first)?,
            tsce::ast_node_array_archive::AstNodeArchive {
                ast_node_type: kind
                    .single_ast_node_type()
                    .expect("single predicate has a comparison node")
                    as i32,
                ..Default::default()
            },
        ]
    };
    Ok(tsce::FormulaArchive {
        ast_node_array: tsce::AstNodeArrayArchive { ast_node: nodes },
        ..Default::default()
    })
}

fn predicate_values(condition: TableCellConditionalHighlightCondition) -> (f64, Option<f64>) {
    if let Some(value) = condition.single_operand() {
        (value.get(), None)
    } else {
        let range = condition.range().expect("range predicate has bounds");
        (range.lower().get(), Some(range.upper().get()))
    }
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

fn range_formula_nodes(
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

fn predicate_number(value: f64) -> Result<tst::FormulaPredArgDataArchive> {
    let decimal = crate::numbers::bnc::decimal128_le(value)?;
    Ok(tst::FormulaPredArgDataArchive {
        double_value: Some(value),
        decimal_low: Some(u64::from_le_bytes(
            decimal[..8]
                .try_into()
                .expect("fixed-size decimal lower half"),
        )),
        decimal_high: Some(u64::from_le_bytes(
            decimal[8..]
                .try_into()
                .expect("fixed-size decimal upper half"),
        )),
        ..Default::default()
    })
}
