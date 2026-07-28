//! Conditional-highlight references stored by table cells.

mod delete;
mod dependencies;
mod formula;
mod native;
mod read;
#[cfg(test)]
mod tests;
mod write;

use chrono::{Days, Local, Months, NaiveDate};
use prost::Message;

use super::*;
use crate::numbers::formula_owner::{formula_owner_uuid_for_table, uuid_as_cfuuid};
use crate::table_cell_conditional_highlight::{
    TableCellConditionalHighlightCondition, TableCellConditionalHighlightDate,
    TableCellConditionalHighlightDateOffsetDirection, TableCellConditionalHighlightDatePeriod,
    TableCellConditionalHighlightDatePeriodUnit, TableCellConditionalHighlightDateRange,
    TableCellConditionalHighlightRule,
};
use native::{
    BINARY_FUNCTION_ARGUMENT_COUNT, BOOLEAN_VALUE_TYPE_CODE, BooleanPredicateKind,
    CELL_DATA_FORMAT_FUNCTION_INDEX, CHECKBOX_DATA_FORMAT_CODE,
    CONDITIONAL_FUNCTION_ARGUMENT_COUNT, CONDITIONAL_FUNCTION_INDEX, CellPredicateKind,
    CheckboxPredicateKind, DATE_DAY_FUNCTION_INDEX, DATE_MONTH_FUNCTION_INDEX,
    DATE_YEAR_FUNCTION_INDEX, FixedDatePredicateKind, IF_ERROR_FUNCTION_INDEX,
    IS_BLANK_FUNCTION_INDEX, IS_ERROR_FUNCTION_INDEX, IS_NUMBER_FUNCTION_INDEX,
    LOGICAL_AND_FUNCTION_INDEX, LOGICAL_NOT_FUNCTION_INDEX, LOGICAL_OR_FUNCTION_INDEX,
    NativePredicateKind, NumericPredicateKind, NumericSignPredicateKind, PREDICATE_ARGUMENT_DATE,
    PREDICATE_ARGUMENT_NONE, PREDICATE_ARGUMENT_NUMBER, PREDICATE_ARGUMENT_RELATIVE_CELL,
    PREDICATE_ARGUMENT_STRING, PREDICATE_CELL_ARGUMENT_INDEX, PREDICATE_DATE_ARGUMENT_INDEX,
    PREDICATE_DATE_EQUALITY_ARGUMENT_INDEX, PREDICATE_DATE_EQUALITY_CELL_ARGUMENT_INDEX,
    PREDICATE_NUMBER_ARGUMENT_INDEX, PREDICATE_QUALIFIER_NONE, PREDICATE_RANGE_CELL_ARGUMENT_INDEX,
    PREDICATE_RANGE_LOWER_ARGUMENT_INDEX, PREDICATE_RANGE_UPPER_ARGUMENT_INDEX,
    PREDICATE_TEXT_ARGUMENT_INDEX, PREDICATE_UNUSED_ARGUMENT_INDEX, RelativeDatePredicateKind,
    TERNARY_FUNCTION_ARGUMENT_COUNT, TEXT_LENGTH_FUNCTION_INDEX, TEXT_RIGHT_FUNCTION_INDEX,
    TEXT_SEARCH_FUNCTION_INDEX, TODAY_FUNCTION_INDEX, TextPredicateKind,
    UNARY_FUNCTION_ARGUMENT_COUNT, VALUE_TYPE_FUNCTION_INDEX, ZERO_FUNCTION_ARGUMENT_COUNT,
};

const MAX_CONDITIONAL_HIGHLIGHT_RULES: usize = CONDITIONAL_STYLE_NO_APPLIED_RULE as usize;
const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const APPLE_EPOCH_YEAR: i32 = 2001;
const APPLE_EPOCH_MONTH: u32 = 1;
const APPLE_EPOCH_DAY: u32 = 1;
const SECONDS_PER_DAY: f64 = 86_400.0;

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

pub(super) fn rules_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Vec<TableCellConditionalHighlightRule>>> {
    let location = locate_cell(package, table_id, row, column)?;
    read::rules_at_location(package, &location, column)
}

pub(super) fn attached_rules_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<Vec<TableCellConditionalHighlightRule>>> {
    let location = locate_attached_cell(package, table_id, row, column)?;
    read::rules_at_location(package, &location, column)
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
    delete::clear_at_location(package, location, row, column)
}

pub(super) fn clear_attached_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<()> {
    let location = locate_attached_cell(package, table_id, row, column)?;
    delete::clear_at_location(package, location, row, column)
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

    let (list_archive, model_archive, conditional_owner_uid) =
        ensure_conditional_style_table(package, &location, list_id)?;
    let style_set_id = first_graph_id;
    let formula_owner_uuid =
        formula_owner_uuid_for_table(&parse_table_uuid(&location.descriptor.model.table_id)?);
    write::insert_conditional_style_graph(
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
    if rules
        .iter()
        .any(|rule| is_volatile_date_condition(&rule.condition))
    {
        dependencies::ensure_volatile_owner(
            package,
            &parse_table_uuid(&location.descriptor.model.table_id)?,
            conditional_owner_uid,
            row,
            column,
        )?;
    }
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
) -> Result<(String, String, tsp::Uuid)> {
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
        let owner_uid = location
            .descriptor
            .model
            .conditional_style_formula_owner_id
            .as_ref()
            .and_then(cfuuid_as_uuid)
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers conditional-style formula owner is missing".to_owned(),
                )
            })?;
        return Ok((archive.clone(), archive.clone(), owner_uid));
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
    let owner_uid = fresh_uuid();
    let owner = uuid_as_cfuuid(&owner_uid);
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
    Ok((model_archive.clone(), model_archive, owner_uid))
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

fn fresh_uuid() -> tsp::Uuid {
    let bytes = litchi_core::id::generate_guid_bytes();
    tsp::Uuid {
        upper: u64::from_be_bytes(bytes[..8].try_into().expect("fixed-size UUID upper half")),
        lower: u64::from_be_bytes(bytes[8..].try_into().expect("fixed-size UUID lower half"))
            .max(3),
    }
}

fn cfuuid_as_uuid(uuid: &tsp::CfuuidArchive) -> Option<tsp::Uuid> {
    let words = || {
        Some(tsp::Uuid {
            lower: u64::from(uuid.uuid_w0?) | (u64::from(uuid.uuid_w1?) << 32),
            upper: u64::from(uuid.uuid_w2?) | (u64::from(uuid.uuid_w3?) << 32),
        })
    };
    let bytes = || {
        let bytes: [u8; 16] = uuid.uuid_bytes.as_deref()?.try_into().ok()?;
        let value = u128::from_be_bytes(bytes);
        Some(tsp::Uuid {
            lower: value as u64,
            upper: (value >> 64) as u64,
        })
    };
    words().or_else(bytes)
}

fn is_volatile_date_condition(condition: &TableCellConditionalHighlightCondition) -> bool {
    matches!(
        condition,
        TableCellConditionalHighlightCondition::DateIsToday
            | TableCellConditionalHighlightCondition::DateIsYesterday
            | TableCellConditionalHighlightCondition::DateIsTomorrow
            | TableCellConditionalHighlightCondition::DateIsInNext(_)
            | TableCellConditionalHighlightCondition::DateIsInLast(_)
            | TableCellConditionalHighlightCondition::DateIsOffsetFromToday(_)
    )
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
    let data = read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?;
    let value = match data.as_deref().map(BncCell::parse).transpose()? {
        None => ConditionalCellValue::Blank,
        Some(cell) if cell.stored_value() == StoredValue::Empty => ConditionalCellValue::Blank,
        Some(cell) => match cell.stored_value() {
            StoredValue::Text(identifier) => {
                let requested = HashSet::from([identifier]);
                let values = resolve_table_string_values(
                    package,
                    &location.object_locations,
                    location
                        .descriptor
                        .model
                        .base_data_store
                        .string_table
                        .identifier,
                    &requested,
                )?;
                let Some(value) = values.into_values().next() else {
                    return Err(Error::InvalidFormat(format!(
                        "iWork conditional-highlight cell references missing string {identifier}"
                    )));
                };
                ConditionalCellValue::Text(value)
            },
            _ => match cell.cached_scalar()? {
                Some(CachedScalar::Number(value)) => ConditionalCellValue::Number(value),
                Some(CachedScalar::Date(value)) => ConditionalCellValue::Date(value),
                Some(CachedScalar::Boolean(value))
                    if cell.cell_format_kind()
                        == Some(crate::numbers::bnc::CHECKBOX_CELL_FORMAT_KIND)
                        && cell.control_cell_spec_identifier().is_some() =>
                {
                    ConditionalCellValue::Checkbox(value)
                },
                Some(CachedScalar::Boolean(value)) => ConditionalCellValue::Boolean(value),
                _ => ConditionalCellValue::Other,
            },
        },
    };
    let date_context = rules
        .iter()
        .any(|rule| {
            matches!(
                rule.condition,
                TableCellConditionalHighlightCondition::DateIsToday
                    | TableCellConditionalHighlightCondition::DateIsYesterday
                    | TableCellConditionalHighlightCondition::DateIsTomorrow
                    | TableCellConditionalHighlightCondition::DateIsInNext(_)
                    | TableCellConditionalHighlightCondition::DateIsInLast(_)
                    | TableCellConditionalHighlightCondition::DateIsOffsetFromToday(_)
            )
        })
        .then(current_date_context);
    rules
        .iter()
        .position(|rule| condition_matches_at(&rule.condition, &value, date_context))
        .map(|index| {
            u32::try_from(index).map_err(|_| {
                Error::ParseError("conditional-highlight rule index exceeds u32".to_owned())
            })
        })
        .transpose()
        .map(|matched| matched.unwrap_or(CONDITIONAL_STYLE_NO_APPLIED_RULE))
}

enum ConditionalCellValue {
    Blank,
    Boolean(bool),
    Checkbox(bool),
    Date(f64),
    Number(f64),
    Other,
    Text(String),
}

#[derive(Clone, Copy)]
struct ConditionalDateContext {
    today: NaiveDate,
    apple_seconds: f64,
}

#[cfg(test)]
fn condition_matches(
    condition: &TableCellConditionalHighlightCondition,
    value: &ConditionalCellValue,
) -> bool {
    condition_matches_at(condition, value, Some(current_date_context()))
}

fn condition_matches_at(
    condition: &TableCellConditionalHighlightCondition,
    value: &ConditionalCellValue,
    date_context: Option<ConditionalDateContext>,
) -> bool {
    match (condition, value) {
        (TableCellConditionalHighlightCondition::CellIsBlank, ConditionalCellValue::Blank) => true,
        (
            TableCellConditionalHighlightCondition::CellIsNotBlank,
            ConditionalCellValue::Number(_)
            | ConditionalCellValue::Boolean(_)
            | ConditionalCellValue::Checkbox(_)
            | ConditionalCellValue::Date(_)
            | ConditionalCellValue::Other
            | ConditionalCellValue::Text(_),
        ) => true,
        (
            TableCellConditionalHighlightCondition::CheckboxIsChecked,
            ConditionalCellValue::Checkbox(value),
        ) => *value,
        (
            TableCellConditionalHighlightCondition::CheckboxIsNotChecked,
            ConditionalCellValue::Checkbox(value),
        ) => !*value,
        (
            TableCellConditionalHighlightCondition::BooleanIsTrue,
            ConditionalCellValue::Boolean(value) | ConditionalCellValue::Checkbox(value),
        ) => *value,
        (
            TableCellConditionalHighlightCondition::BooleanIsFalse,
            ConditionalCellValue::Boolean(value) | ConditionalCellValue::Checkbox(value),
        ) => !*value,
        (
            TableCellConditionalHighlightCondition::NumberIsPositive,
            ConditionalCellValue::Number(value),
        ) => *value > 0.0,
        (
            TableCellConditionalHighlightCondition::NumberIsNegative,
            ConditionalCellValue::Number(value),
        ) => *value < 0.0,
        (
            TableCellConditionalHighlightCondition::DateIsToday,
            ConditionalCellValue::Date(value),
        ) => date_context.is_some_and(|context| {
            *value >= context.apple_seconds && *value < context.apple_seconds + SECONDS_PER_DAY
        }),
        (
            TableCellConditionalHighlightCondition::DateIsYesterday,
            ConditionalCellValue::Date(value),
        ) => date_context.is_some_and(|context| {
            *value >= context.apple_seconds - SECONDS_PER_DAY && *value < context.apple_seconds
        }),
        (
            TableCellConditionalHighlightCondition::DateIsTomorrow,
            ConditionalCellValue::Date(value),
        ) => date_context.is_some_and(|context| {
            *value >= context.apple_seconds + SECONDS_PER_DAY
                && *value < context.apple_seconds + 2.0 * SECONDS_PER_DAY
        }),
        (
            TableCellConditionalHighlightCondition::DateIsInNext(period),
            ConditionalCellValue::Date(value),
        ) => date_context.is_some_and(|context| {
            shifted_date(
                context.today,
                *period,
                TableCellConditionalHighlightDateOffsetDirection::FromNow,
            )
            .is_some_and(|upper| {
                *value >= context.apple_seconds
                    && *value < date_to_apple_seconds(upper) + SECONDS_PER_DAY
            })
        }),
        (
            TableCellConditionalHighlightCondition::DateIsInLast(period),
            ConditionalCellValue::Date(value),
        ) => date_context.is_some_and(|context| {
            shifted_date(
                context.today,
                *period,
                TableCellConditionalHighlightDateOffsetDirection::Ago,
            )
            .is_some_and(|lower| {
                *value >= date_to_apple_seconds(lower)
                    && *value < context.apple_seconds + SECONDS_PER_DAY
            })
        }),
        (
            TableCellConditionalHighlightCondition::DateIsOffsetFromToday(offset),
            ConditionalCellValue::Date(value),
        ) => date_context.is_some_and(|context| {
            shifted_date(context.today, offset.period(), offset.direction()).is_some_and(|target| {
                let lower = date_to_apple_seconds(target);
                *value >= lower && *value < lower + SECONDS_PER_DAY
            })
        }),
        (
            TableCellConditionalHighlightCondition::DateIs(date),
            ConditionalCellValue::Date(value),
        ) => {
            let lower = date.apple_seconds();
            *value >= lower && *value < lower + SECONDS_PER_DAY
        },
        (
            TableCellConditionalHighlightCondition::DateIsBefore(date),
            ConditionalCellValue::Date(value),
        ) => *value < date.apple_seconds(),
        (
            TableCellConditionalHighlightCondition::DateIsAfter(date),
            ConditionalCellValue::Date(value),
        ) => *value >= date.apple_seconds() + SECONDS_PER_DAY,
        (
            TableCellConditionalHighlightCondition::DateIsBetween(range),
            ConditionalCellValue::Date(value),
        ) => {
            *value >= range.lower().apple_seconds()
                && *value < range.upper().apple_seconds() + SECONDS_PER_DAY
        },
        (
            TableCellConditionalHighlightCondition::EqualTo(operand),
            ConditionalCellValue::Number(value),
        ) => *value == operand.get(),
        (
            TableCellConditionalHighlightCondition::NotEqualTo(operand),
            ConditionalCellValue::Number(value),
        ) => *value != operand.get(),
        (
            TableCellConditionalHighlightCondition::GreaterThan(operand),
            ConditionalCellValue::Number(value),
        ) => *value > operand.get(),
        (
            TableCellConditionalHighlightCondition::GreaterThanOrEqualTo(operand),
            ConditionalCellValue::Number(value),
        ) => *value >= operand.get(),
        (
            TableCellConditionalHighlightCondition::LessThan(operand),
            ConditionalCellValue::Number(value),
        ) => *value < operand.get(),
        (
            TableCellConditionalHighlightCondition::LessThanOrEqualTo(operand),
            ConditionalCellValue::Number(value),
        ) => *value <= operand.get(),
        (
            TableCellConditionalHighlightCondition::Between(range),
            ConditionalCellValue::Number(value),
        ) => *value >= range.lower().get() && *value <= range.upper().get(),
        (
            TableCellConditionalHighlightCondition::NotBetween(range),
            ConditionalCellValue::Number(value),
        ) => *value < range.lower().get() || *value > range.upper().get(),
        (
            TableCellConditionalHighlightCondition::TextEqualTo(needle),
            ConditionalCellValue::Text(value),
        ) => equals_case_insensitive(value, needle.as_str()),
        (
            TableCellConditionalHighlightCondition::TextNotEqualTo(needle),
            ConditionalCellValue::Text(value),
        ) => !equals_case_insensitive(value, needle.as_str()),
        (
            TableCellConditionalHighlightCondition::TextStartsWith(needle),
            ConditionalCellValue::Text(value),
        ) => starts_with_case_insensitive(value, needle.as_str()),
        (
            TableCellConditionalHighlightCondition::TextDoesNotStartWith(needle),
            ConditionalCellValue::Text(value),
        ) => !starts_with_case_insensitive(value, needle.as_str()),
        (
            TableCellConditionalHighlightCondition::TextEndsWith(needle),
            ConditionalCellValue::Text(value),
        ) => ends_with_case_insensitive(value, needle.as_str()),
        (
            TableCellConditionalHighlightCondition::TextDoesNotEndWith(needle),
            ConditionalCellValue::Text(value),
        ) => !ends_with_case_insensitive(value, needle.as_str()),
        (
            TableCellConditionalHighlightCondition::TextContains(needle),
            ConditionalCellValue::Text(value),
        ) => contains_case_insensitive(value, needle.as_str()),
        (
            TableCellConditionalHighlightCondition::TextDoesNotContain(needle),
            ConditionalCellValue::Text(value),
        ) => !contains_case_insensitive(value, needle.as_str()),
        _ => false,
    }
}

fn current_date_context() -> ConditionalDateContext {
    let today = Local::now().date_naive();
    ConditionalDateContext {
        today,
        apple_seconds: date_to_apple_seconds(today),
    }
}

fn date_to_apple_seconds(date: NaiveDate) -> f64 {
    let epoch = NaiveDate::from_ymd_opt(APPLE_EPOCH_YEAR, APPLE_EPOCH_MONTH, APPLE_EPOCH_DAY)
        .expect("the Apple epoch is a valid calendar date");
    date.signed_duration_since(epoch).num_days() as f64 * SECONDS_PER_DAY
}

fn shifted_date(
    date: NaiveDate,
    period: TableCellConditionalHighlightDatePeriod,
    direction: TableCellConditionalHighlightDateOffsetDirection,
) -> Option<NaiveDate> {
    let forward = matches!(
        direction,
        TableCellConditionalHighlightDateOffsetDirection::FromNow
    );
    match period.unit() {
        TableCellConditionalHighlightDatePeriodUnit::Days => {
            shift_days(date, u64::from(period.count()), forward)
        },
        TableCellConditionalHighlightDatePeriodUnit::Weeks => period
            .count()
            .checked_mul(7)
            .and_then(|days| shift_days(date, u64::from(days), forward)),
        TableCellConditionalHighlightDatePeriodUnit::Months => {
            shift_months(date, period.count(), forward)
        },
        TableCellConditionalHighlightDatePeriodUnit::Quarters => period
            .count()
            .checked_mul(3)
            .and_then(|months| shift_months(date, months, forward)),
        TableCellConditionalHighlightDatePeriodUnit::Years => period
            .count()
            .checked_mul(12)
            .and_then(|months| shift_months(date, months, forward)),
    }
}

fn shift_days(date: NaiveDate, days: u64, forward: bool) -> Option<NaiveDate> {
    if forward {
        date.checked_add_days(Days::new(days))
    } else {
        date.checked_sub_days(Days::new(days))
    }
}

fn shift_months(date: NaiveDate, months: u32, forward: bool) -> Option<NaiveDate> {
    if forward {
        date.checked_add_months(Months::new(months))
    } else {
        date.checked_sub_months(Months::new(months))
    }
}

fn equals_case_insensitive(value: &str, expected: &str) -> bool {
    if value.is_ascii() && expected.is_ascii() {
        return value.eq_ignore_ascii_case(expected);
    }
    value.to_lowercase() == expected.to_lowercase()
}

fn starts_with_case_insensitive(value: &str, prefix: &str) -> bool {
    if value.is_ascii() && prefix.is_ascii() {
        return value
            .as_bytes()
            .get(..prefix.len())
            .is_some_and(|candidate| candidate.eq_ignore_ascii_case(prefix.as_bytes()));
    }
    value.to_lowercase().starts_with(&prefix.to_lowercase())
}

fn ends_with_case_insensitive(value: &str, suffix: &str) -> bool {
    if value.is_ascii() && suffix.is_ascii() {
        return value
            .as_bytes()
            .get(value.len().saturating_sub(suffix.len())..)
            .is_some_and(|candidate| {
                candidate.len() == suffix.len() && candidate.eq_ignore_ascii_case(suffix.as_bytes())
            });
    }
    value.to_lowercase().ends_with(&suffix.to_lowercase())
}

fn contains_case_insensitive(value: &str, needle: &str) -> bool {
    if value.is_ascii() && needle.is_ascii() {
        return value
            .as_bytes()
            .windows(needle.len())
            .any(|candidate| candidate.eq_ignore_ascii_case(needle.as_bytes()));
    }
    value.to_lowercase().contains(&needle.to_lowercase())
}
