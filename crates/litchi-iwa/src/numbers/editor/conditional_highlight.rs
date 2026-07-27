//! Conditional-highlight references stored by table cells.

use prost::Message;

use super::*;
use crate::numbers::formula_owner::{formula_owner_uuid_for_table, uuid_as_cfuuid};
use crate::table_cell_conditional_highlight::{
    TableCellConditionalHighlightCondition, TableCellConditionalHighlightRule,
};

const MAX_CONDITIONAL_HIGHLIGHT_RULES: usize = CONDITIONAL_STYLE_NO_APPLIED_RULE as usize;
const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const CONDITIONAL_STYLE_SET_MESSAGE_TYPE: u32 = 6_010;
const CELL_STYLE_MESSAGE_TYPE: u32 = 6_004;
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const NATIVE_MESSAGE_VERSION: &[u32] = &[1, 0, 5];

pub(super) fn info_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellConditionalHighlightInfo>> {
    let location = locate_cell(package, table_id, row, column)?;
    info_at_location(package, location, table_id, row, column)
}

pub(super) fn attached_info_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellConditionalHighlightInfo>> {
    let location = locate_attached_cell(package, table_id, row, column)?;
    info_at_location(package, location, table_id, row, column)
}

fn info_at_location(
    package: &IWorkPackage,
    location: CellLocation,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellConditionalHighlightInfo>> {
    let Some(cell) = read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    else {
        return Ok(None);
    };
    let bnc = BncCell::parse(&cell)?;
    let Some(list_identifier) = bnc.conditional_style_identifier() else {
        return Ok(None);
    };
    if bnc
        .conditional_style_applied_rule()
        .is_some_and(|rule| rule > CONDITIONAL_STYLE_NO_APPLIED_RULE)
    {
        return Err(Error::InvalidFormat(
            "iWork conditional-highlight applied-rule index is out of range".to_owned(),
        ));
    }
    let (_resolved, entry) = resolve_entry(package, &location, list_identifier)?;
    let style_set_object_id = entry
        .entry
        .reference
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-highlight entry {list_identifier} has no style-set reference"
            ))
        })?;
    let archive_name = location
        .object_locations
        .get(&style_set_object_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-highlight style set {style_set_object_id} is missing"
            ))
        })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(style_set_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork conditional-highlight style set {style_set_object_id} is missing"
        ))
    })?;
    let style_set = object
        .messages
        .iter()
        .find_map(|message| {
            (message.type_ == 6_010)
                .then(|| tst::ConditionalStyleSetArchive::decode(message.data.as_slice()))
        })
        .transpose()?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-highlight object {style_set_object_id} has no style-set payload"
            ))
        })?;
    Ok(Some(TableCellConditionalHighlightInfo {
        table_id,
        row,
        column,
        list_identifier,
        style_set_object_id,
        rule_count: style_set.rule_count,
    }))
}

pub(super) fn clear_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<()> {
    let location = locate_cell(package, table_id, row, column)?;
    clear_at_location(package, location, row, column)
}

pub(super) fn clear_attached_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<()> {
    let location = locate_attached_cell(package, table_id, row, column)?;
    clear_at_location(package, location, row, column)
}

pub(super) fn set_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    rules: &[TableCellConditionalHighlightRule],
) -> Result<()> {
    validate_rules(rules)?;
    clear_in_package(package, table_id, row, column)?;
    let location = locate_cell(package, table_id, row, column)?;
    set_at_location(package, location, row, column, rules)
}

pub(super) fn set_attached_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    rules: &[TableCellConditionalHighlightRule],
) -> Result<()> {
    validate_rules(rules)?;
    clear_attached_in_package(package, table_id, row, column)?;
    let location = locate_attached_cell(package, table_id, row, column)?;
    set_at_location(package, location, row, column, rules)
}

fn validate_rules(rules: &[TableCellConditionalHighlightRule]) -> Result<()> {
    if rules.is_empty() {
        return Err(Error::ParseError(
            "conditional highlighting requires at least one rule".to_owned(),
        ));
    }
    if rules.len() > MAX_CONDITIONAL_HIGHLIGHT_RULES {
        return Err(Error::ParseError(format!(
            "conditional highlighting supports at most {MAX_CONDITIONAL_HIGHLIGHT_RULES} rules"
        )));
    }
    Ok(())
}

fn set_at_location(
    package: &mut IWorkPackage,
    location: CellLocation,
    row: usize,
    column: usize,
    rules: &[TableCellConditionalHighlightRule],
) -> Result<()> {
    let list_id = location
        .descriptor
        .model
        .base_data_store
        .conditionalstyletable
        .as_ref()
        .map(|reference| reference.identifier);
    let first_id = next_object_identifier(package)?;
    let list_id = list_id.unwrap_or(first_id);
    let first_graph_id = if list_id == first_id {
        first_id.checked_add(1)
    } else {
        Some(first_id)
    }
    .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
    let graph_id_count = 1usize
        .checked_add(rules.len().checked_mul(2).ok_or_else(|| {
            Error::ParseError("conditional-highlight object count overflow".to_owned())
        })?)
        .ok_or_else(|| {
            Error::ParseError("conditional-highlight object count overflow".to_owned())
        })?;
    let last_graph_id = first_graph_id
        .checked_add(u64::try_from(graph_id_count - 1).map_err(|_| {
            Error::ParseError("conditional-highlight object count exceeds u64".to_owned())
        })?)
        .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;

    let (list_archive, model_archive) =
        ensure_conditional_style_table(package, &location, list_id)?;
    let style_set_id = first_graph_id;
    let formula_owner_uuid =
        formula_owner_uuid_for_table(&parse_table_uuid(&location.descriptor.model.table_id)?);
    insert_conditional_style_graph(
        package,
        &list_archive,
        style_set_id,
        rules,
        &formula_owner_uuid,
    )?;
    let list_identifier =
        append_conditional_style_entry(package, &list_archive, list_id, style_set_id)?;

    let component = component_identifier_for_entry(package, &model_archive)?;
    if let Some(component) = component {
        let mut ids = Vec::with_capacity(graph_id_count + usize::from(list_id == first_id));
        if list_id == first_id {
            ids.push(list_id);
        }
        ids.extend(first_graph_id..=last_graph_id);
        add_component_object_uuids(package, component, &ids)?;
    }
    set_package_last_object_identifier(package, last_graph_id.max(list_id))?;
    let applied_rule = applied_rule_for_cell(package, &location, column, rules)?;
    update_cell(
        package,
        &location,
        row,
        column,
        Some(list_identifier),
        Some(applied_rule),
    )?;
    let mut modified = vec![location.tile_archive, list_archive];
    if !modified.contains(&model_archive) {
        modified.push(model_archive);
    }
    advance_save_tokens_for_entries(package, &modified)
}

fn ensure_conditional_style_table(
    package: &mut IWorkPackage,
    location: &CellLocation,
    list_id: u64,
) -> Result<(String, String)> {
    if let Some(reference) = &location
        .descriptor
        .model
        .base_data_store
        .conditionalstyletable
    {
        let archive = location
            .object_locations
            .get(&reference.identifier)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers conditional-style table object {} is missing",
                    reference.identifier
                ))
            })?;
        return Ok((archive.clone(), archive.clone()));
    }
    let model_archive = location
        .object_locations
        .get(&location.descriptor.object_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table object {} is missing",
                location.descriptor.object_id
            ))
        })?
        .clone();
    let owner = fresh_cfuuid();
    package.update_archive(&model_archive, |archive| {
        archive.insert_object(ArchiveObject::new(
            list_id,
            vec![RawMessage {
                type_: TABLE_DATA_LIST_MESSAGE_TYPE,
                data: TableDataList {
                    list_type: tst::table_data_list::ListType::ConditionalStyle as i32,
                    next_list_id: 1,
                    entries: Vec::new(),
                    segments: Vec::new(),
                    is_new_for_bnc: Some(true),
                }
                .encode_to_vec(),
            }],
        )?)?;
        let object = archive
            .object_mut(location.descriptor.object_id)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers table object {} is missing",
                    location.descriptor.object_id
                ))
            })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == 6_000 || message.type_ == 6_001)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {} has no Numbers table-model payload",
                    location.descriptor.object_id
                ))
            })?;
        let original = object.messages[message_index].data.as_slice();
        let previous = TableModelArchive::decode(original)?;
        let mut model = previous.clone();
        model.base_data_store.conditionalstyletable = Some(tsp::Reference {
            identifier: list_id,
            ..Default::default()
        });
        model.conditional_style_formula_owner_id = Some(owner.clone());
        let data = rewrite_table_model_conditional_style_wire(original, &previous, &model)?;
        let message_type = object.messages[message_index].type_;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        add_message_object_reference(object, message_index, list_id, list_id);
        Ok(())
    })?;
    Ok((model_archive.clone(), model_archive))
}

fn append_conditional_style_entry(
    package: &mut IWorkPackage,
    archive_name: &str,
    list_id: u64,
    style_set_id: u64,
) -> Result<u32> {
    let locations = object_locations(package)?;
    let resolved = resolve_table_data_list(
        package,
        &locations,
        list_id,
        tst::table_data_list::ListType::ConditionalStyle,
    )?;
    let key = next_table_data_list_key(&resolved.list, &resolved.entries)?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(list_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers conditional-style table object {list_id} is missing"
            ))
        })?;
        let message_index =
            table_data_list_message_index(object, tst::table_data_list::ListType::ConditionalStyle)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Object {list_id} has no conditional-style TableDataList payload"
                    ))
                })?;
        let original = object.messages[message_index].data.as_slice();
        let previous = TableDataList::decode(original)?;
        let mut list = previous.clone();
        list.next_list_id = key.checked_add(1).ok_or_else(|| {
            Error::ParseError("conditional-style list identifier overflow".to_owned())
        })?;
        list.entries.push(tst::table_data_list::ListEntry {
            key,
            refcount: 1,
            reference: Some(tsp::Reference {
                identifier: style_set_id,
                ..Default::default()
            }),
            ..Default::default()
        });
        let data = rewrite_table_data_list_wire(original, &previous, &list)?;
        let message_type = object.messages[message_index].type_;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        add_message_object_reference(object, message_index, style_set_id, style_set_id);
        Ok(())
    })?;
    Ok(key)
}

fn insert_conditional_style_graph(
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
        let formula = predicate_formula(rule.condition, formula_owner_uuid)?;
        let predicate_type = predicate_type(rule.condition);
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
                    qualifier1: 0,
                    qualifier2: 0,
                    param_index1: 1,
                    param_index2: -1,
                    param_index0: 0,
                },
                cell_style,
                text_style,
            },
        );
        current.push(tst::conditional_style_set_archive::ConditionalStyleRule {
            predicate: Some(tst::FormulaPredicateArchive {
                predicate_type,
                qualifier1: 0,
                qualifier2: 0,
                param_value0: Some(tst::FormulaPredArgArchive {
                    arg_type: 4,
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
                    arg_type: 1,
                    arg_value: Some(predicate_number(rule.condition)?),
                    ..Default::default()
                }),
                param_value2: Some(tst::FormulaPredArgArchive {
                    arg_type: 0,
                    ..Default::default()
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

fn predicate_type(condition: TableCellConditionalHighlightCondition) -> i32 {
    match condition {
        TableCellConditionalHighlightCondition::EqualTo(_) => 5,
        TableCellConditionalHighlightCondition::NotEqualTo(_) => 6,
        TableCellConditionalHighlightCondition::GreaterThan(_) => 7,
        TableCellConditionalHighlightCondition::GreaterThanOrEqualTo(_) => 8,
        TableCellConditionalHighlightCondition::LessThan(_) => 9,
        TableCellConditionalHighlightCondition::LessThanOrEqualTo(_) => 10,
    }
}

fn predicate_formula(
    condition: TableCellConditionalHighlightCondition,
    formula_owner_uuid: &tsp::Uuid,
) -> Result<tsce::FormulaArchive> {
    use tsce::ast_node_array_archive::AstNodeType;
    let comparison = match condition {
        TableCellConditionalHighlightCondition::EqualTo(_) => AstNodeType::EqualToNode,
        TableCellConditionalHighlightCondition::NotEqualTo(_) => AstNodeType::NotEqualToNode,
        TableCellConditionalHighlightCondition::GreaterThan(_) => AstNodeType::GreaterThanNode,
        TableCellConditionalHighlightCondition::GreaterThanOrEqualTo(_) => {
            AstNodeType::GreaterThanOrEqualToNode
        },
        TableCellConditionalHighlightCondition::LessThan(_) => AstNodeType::LessThanNode,
        TableCellConditionalHighlightCondition::LessThanOrEqualTo(_) => {
            AstNodeType::LessThanOrEqualToNode
        },
    };
    let value = condition.operand().get();
    let decimal = crate::numbers::bnc::decimal128_le(value)?;
    let linked_cell = tsce::ast_node_array_archive::AstNodeArchive {
        ast_node_type: AstNodeType::LinkedCellRefNode as i32,
        ast_cross_table_reference_extra_info: Some(
            tsce::ast_node_array_archive::AstCrossTableReferenceExtraInfoArchive {
                table_id: uuid_as_cfuuid(formula_owner_uuid),
                ..Default::default()
            },
        ),
        ..Default::default()
    };
    let number = tsce::ast_node_array_archive::AstNodeArchive {
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
    };
    Ok(tsce::FormulaArchive {
        ast_node_array: tsce::AstNodeArrayArchive {
            ast_node: vec![
                linked_cell,
                number,
                tsce::ast_node_array_archive::AstNodeArchive {
                    ast_node_type: comparison as i32,
                    ..Default::default()
                },
            ],
        },
        ..Default::default()
    })
}

fn predicate_number(
    condition: TableCellConditionalHighlightCondition,
) -> Result<tst::FormulaPredArgDataArchive> {
    let value = condition.operand().get();
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

fn parse_table_uuid(value: &str) -> Result<tsp::Uuid> {
    let compact = value.replace('-', "");
    if compact.len() != 32 {
        return Err(Error::InvalidFormat(format!(
            "invalid Numbers table UUID {value:?}"
        )));
    }
    let bytes = (0..16)
        .map(|index| u8::from_str_radix(&compact[index * 2..index * 2 + 2], 16))
        .collect::<std::result::Result<Vec<_>, _>>()
        .map_err(|_| Error::InvalidFormat(format!("invalid Numbers table UUID {value:?}")))?;
    Ok(tsp::Uuid {
        upper: u64::from_be_bytes(bytes[..8].try_into().expect("fixed-size UUID upper half")),
        lower: u64::from_be_bytes(bytes[8..].try_into().expect("fixed-size UUID lower half")),
    })
}

fn fresh_cfuuid() -> tsp::CfuuidArchive {
    let bytes = litchi_core::id::generate_guid_bytes();
    uuid_as_cfuuid(&tsp::Uuid {
        upper: u64::from_be_bytes(bytes[..8].try_into().expect("fixed-size UUID upper half")),
        lower: u64::from_be_bytes(bytes[8..].try_into().expect("fixed-size UUID lower half")),
    })
}

fn clear_at_location(
    package: &mut IWorkPackage,
    location: CellLocation,
    row: usize,
    column: usize,
) -> Result<()> {
    let Some(cell) = read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    else {
        return Ok(());
    };
    let Some(list_identifier) = BncCell::parse(&cell)?.conditional_style_identifier() else {
        return Ok(());
    };
    let locations = location.object_locations.clone();
    let (resolved, entry) = resolve_entry(package, &location, list_identifier)?;
    let style_set_object_id = entry
        .entry
        .reference
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-highlight entry {list_identifier} has no style-set reference"
            ))
        })?;
    let mut owned_object_ids =
        conditional_style_owned_object_ids(package, &locations, style_set_object_id)?;
    owned_object_ids.push(style_set_object_id);
    let removed = decrement_table_data_list_entry(
        package,
        &locations,
        &resolved,
        &entry,
        tst::table_data_list::ListType::ConditionalStyle,
    )?;
    update_cell(package, &location, row, column, None, None)?;
    let mut modified_entries = vec![location.tile_archive.clone(), resolved.table_archive];
    if removed {
        let style_archive = locations.get(&style_set_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-highlight style set {style_set_object_id} is missing"
            ))
        })?;
        if let Some(component) = component_identifier_for_entry(package, style_archive)? {
            for identifier in &owned_object_ids {
                remove_component_external_references_to_object(package, component, *identifier)?;
            }
            remove_component_object_uuids(package, component, &owned_object_ids)?;
        }
        for identifier in &owned_object_ids {
            if let Some(archive_name) = locations.get(identifier)
                && !modified_entries.contains(archive_name)
            {
                modified_entries.push(archive_name.clone());
            }
            remove_object_or_empty_entry(package, &locations, *identifier)?;
        }
        release_package_identifier_suffix(package, &owned_object_ids)?;
    }
    advance_save_tokens_for_entries(package, &modified_entries)
}

fn conditional_style_owned_object_ids(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    style_set_id: u64,
) -> Result<Vec<u64>> {
    let archive_name = locations.get(&style_set_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork conditional-highlight style set {style_set_id} is missing"
        ))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(style_set_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork conditional-highlight style set {style_set_id} is missing"
        ))
    })?;
    let set = object
        .messages
        .iter()
        .find_map(|message| {
            (message.type_ == CONDITIONAL_STYLE_SET_MESSAGE_TYPE)
                .then(|| tst::ConditionalStyleSetArchive::decode(message.data.as_slice()))
        })
        .transpose()?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-highlight object {style_set_id} has no style-set payload"
            ))
        })?;
    let mut identifiers = Vec::with_capacity(set.rules_prepivot.len() * 2);
    for rule in set.rules_prepivot {
        for identifier in [rule.cell_style.identifier, rule.text_style.identifier] {
            if identifier != 0 && !identifiers.contains(&identifier) {
                identifiers.push(identifier);
            }
        }
    }
    if let Some(rules) = set.rules {
        for rule in rules.rule {
            for identifier in [rule.cell_style.identifier, rule.text_style.identifier] {
                if identifier != 0 && !identifiers.contains(&identifier) {
                    identifiers.push(identifier);
                }
            }
        }
    }
    Ok(identifiers)
}

fn resolve_entry(
    package: &IWorkPackage,
    location: &CellLocation,
    list_identifier: u32,
) -> Result<(ResolvedTableDataList, LocatedTableDataListEntry)> {
    let table_id = location
        .descriptor
        .model
        .base_data_store
        .conditionalstyletable
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork cell references conditional-highlight entry {list_identifier}, but its table has no conditional-style list"
            ))
        })?;
    let resolved = resolve_table_data_list(
        package,
        &location.object_locations,
        table_id,
        tst::table_data_list::ListType::ConditionalStyle,
    )?;
    let entry = resolved
        .entries
        .iter()
        .find(|entry| entry.entry.key == list_identifier)
        .cloned()
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-style list has no entry {list_identifier}"
            ))
        })?;
    Ok((resolved, entry))
}

fn update_cell(
    package: &mut IWorkPackage,
    location: &CellLocation,
    row: usize,
    column: usize,
    identifier: Option<u32>,
    applied_rule: Option<u32>,
) -> Result<()> {
    let cell_count = update_tile(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
        location.descriptor.model.number_of_columns as usize,
        EncodedValue::ConditionalStyle {
            identifier,
            applied_rule,
        },
    )?;
    update_row_header(
        package,
        &location.object_locations,
        &location.descriptor.model,
        row,
        cell_count,
    )
}

fn applied_rule_for_cell(
    package: &IWorkPackage,
    location: &CellLocation,
    column: usize,
    rules: &[TableCellConditionalHighlightRule],
) -> Result<u32> {
    let scalar = read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    .map(|data| BncCell::parse(&data))
    .transpose()?
    .map(|cell| cell.cached_scalar())
    .transpose()?
    .flatten();
    let Some(CachedScalar::Number(value)) = scalar else {
        return Ok(CONDITIONAL_STYLE_NO_APPLIED_RULE);
    };
    rules
        .iter()
        .position(|rule| condition_matches(rule.condition, value))
        .map(|index| {
            u32::try_from(index).map_err(|_| {
                Error::ParseError("conditional-highlight rule index exceeds u32".to_owned())
            })
        })
        .transpose()
        .map(|matched| matched.unwrap_or(CONDITIONAL_STYLE_NO_APPLIED_RULE))
}

fn condition_matches(condition: TableCellConditionalHighlightCondition, value: f64) -> bool {
    let operand = condition.operand().get();
    match condition {
        TableCellConditionalHighlightCondition::EqualTo(_) => value == operand,
        TableCellConditionalHighlightCondition::NotEqualTo(_) => value != operand,
        TableCellConditionalHighlightCondition::GreaterThan(_) => value > operand,
        TableCellConditionalHighlightCondition::GreaterThanOrEqualTo(_) => value >= operand,
        TableCellConditionalHighlightCondition::LessThan(_) => value < operand,
        TableCellConditionalHighlightCondition::LessThanOrEqualTo(_) => value <= operand,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::{RgbColorSpace, RgbaColor};
    use crate::table_cell_conditional_highlight::{
        TableCellConditionalHighlightNumber, TableCellConditionalHighlightStyle,
    };

    fn rule(
        condition: TableCellConditionalHighlightCondition,
        color: RgbaColor,
    ) -> TableCellConditionalHighlightRule {
        TableCellConditionalHighlightRule::new(
            condition,
            TableCellConditionalHighlightStyle::with_fill(color),
        )
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
        let zero = TableCellConditionalHighlightNumber::new(0.0).unwrap();
        let hundred = TableCellConditionalHighlightNumber::new(100.0).unwrap();
        let initial = [
            rule(TableCellConditionalHighlightCondition::LessThan(zero), red),
            rule(
                TableCellConditionalHighlightCondition::GreaterThanOrEqualTo(hundred),
                green,
            ),
        ];

        editor
            .set_cell_conditional_highlighting(table_id, 1, 1, &initial)
            .unwrap();
        assert_eq!(
            editor
                .cell_conditional_highlighting(table_id, 1, 1)
                .unwrap()
                .unwrap()
                .rule_count,
            2
        );

        editor
            .set_cell_conditional_highlighting(table_id, 1, 1, &initial[..1])
            .unwrap();
        assert_eq!(
            editor
                .cell_conditional_highlighting(table_id, 1, 1)
                .unwrap()
                .unwrap()
                .rule_count,
            1
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
}
