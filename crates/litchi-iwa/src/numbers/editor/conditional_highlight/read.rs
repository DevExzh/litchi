//! Strict semantic decoding of native conditional-highlight rule graphs.

use prost::Message;

use super::*;
use crate::table_cell_conditional_highlight::{
    TableCellConditionalHighlightNumber, TableCellConditionalHighlightRule,
    TableCellConditionalHighlightStyle,
};

const CONDITIONAL_STYLE_SET_MESSAGE_TYPE: u32 = 6_010;
const CELL_STYLE_MESSAGE_TYPE: u32 = 6_004;
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;

pub(super) fn rules_at_location(
    package: &IWorkPackage,
    location: &CellLocation,
    column: usize,
) -> Result<Option<Vec<TableCellConditionalHighlightRule>>> {
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
    let Some(list_identifier) = BncCell::parse(&cell)?.conditional_style_identifier() else {
        return Ok(None);
    };
    let (_resolved, entry) = resolve_entry(package, location, list_identifier)?;
    let style_set_id = entry
        .entry
        .reference
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-highlight entry {list_identifier} has no style-set reference"
            ))
        })?;
    let style_set = decode_style_set(package, &location.object_locations, style_set_id)?;
    decode_rules(package, &location.object_locations, &style_set).map(Some)
}

fn decode_style_set(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    style_set_id: u64,
) -> Result<tst::ConditionalStyleSetArchive> {
    with_object(package, locations, style_set_id, "style set", |object| {
        object
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
            })
    })
}

fn decode_rules(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    set: &tst::ConditionalStyleSetArchive,
) -> Result<Vec<TableCellConditionalHighlightRule>> {
    let current = set.rules.as_ref().ok_or_else(unsupported_rule_graph)?;
    let declared = usize::try_from(set.rule_count).map_err(|_| {
        Error::InvalidFormat("conditional-highlight rule count overflow".to_owned())
    })?;
    if current.rule.len() != declared || set.rules_prepivot.len() != declared {
        return Err(Error::InvalidFormat(
            "conditional-highlight rule representations disagree on their count".to_owned(),
        ));
    }
    current
        .rule
        .iter()
        .zip(&set.rules_prepivot)
        .map(|(rule, prepivot)| decode_rule(package, locations, rule, prepivot))
        .collect()
}

fn decode_rule(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    rule: &tst::conditional_style_set_archive::ConditionalStyleRule,
    prepivot: &tst::conditional_style_set_archive::ConditionalStyleRulePrePivot,
) -> Result<TableCellConditionalHighlightRule> {
    if rule.cell_style.identifier != prepivot.cell_style.identifier
        || rule.text_style.identifier != prepivot.text_style.identifier
    {
        return Err(unsupported_rule_graph());
    }
    let predicate = rule.predicate.as_ref().ok_or_else(unsupported_rule_graph)?;
    let value = decode_predicate(predicate, &prepivot.predicate)?;
    let condition = NumericPredicateKind::try_from(predicate.predicate_type)
        .map_err(|_| unsupported_rule_graph())?
        .condition(value);
    let fill = decode_cell_style(package, locations, rule.cell_style.identifier)?;
    let (text_color, bold) = decode_text_style(package, locations, rule.text_style.identifier)?;
    Ok(TableCellConditionalHighlightRule::new(
        condition,
        TableCellConditionalHighlightStyle::new(fill, text_color, bold).map_err(|_| {
            Error::InvalidFormat(
                "conditional-highlight rule has no supported visual override".to_owned(),
            )
        })?,
    ))
}

fn decode_predicate(
    predicate: &tst::FormulaPredicateArchive,
    prepivot: &tst::FormulaPredicatePrePivotArchive,
) -> Result<TableCellConditionalHighlightNumber> {
    if predicate.predicate_type != prepivot.predicate_type
        || predicate.qualifier1 != PREDICATE_QUALIFIER_NONE
        || predicate.qualifier2 != PREDICATE_QUALIFIER_NONE
        || prepivot.qualifier1 != PREDICATE_QUALIFIER_NONE
        || prepivot.qualifier2 != PREDICATE_QUALIFIER_NONE
        || prepivot.param_index0 != PREDICATE_CELL_ARGUMENT_INDEX
        || prepivot.param_index1 != PREDICATE_NUMBER_ARGUMENT_INDEX
        || prepivot.param_index2 != PREDICATE_UNUSED_ARGUMENT_INDEX
        || predicate.for_conditional_style != Some(true)
        || predicate
            .param_value0
            .as_ref()
            .map(|argument| argument.arg_type)
            != Some(PREDICATE_ARGUMENT_RELATIVE_CELL)
        || predicate
            .param_value2
            .as_ref()
            .map(|argument| argument.arg_type)
            != Some(PREDICATE_ARGUMENT_NONE)
    {
        return Err(unsupported_rule_graph());
    }
    let argument = predicate
        .param_value1
        .as_ref()
        .filter(|argument| argument.arg_type == PREDICATE_ARGUMENT_NUMBER)
        .and_then(|argument| argument.arg_value.as_ref())
        .ok_or_else(unsupported_rule_graph)?;
    let value = argument.double_value.ok_or_else(unsupported_rule_graph)?;
    let number =
        TableCellConditionalHighlightNumber::new(value).map_err(|_| unsupported_rule_graph())?;
    validate_formula(
        predicate
            .formula
            .as_ref()
            .ok_or_else(unsupported_rule_graph)?,
        predicate.predicate_type,
        value,
    )?;
    validate_formula(&prepivot.formula, prepivot.predicate_type, value)?;
    Ok(number)
}

fn validate_formula(formula: &tsce::FormulaArchive, predicate_type: i32, value: f64) -> Result<()> {
    use tsce::ast_node_array_archive::AstNodeType;

    let comparison = NumericPredicateKind::try_from(predicate_type)
        .map_err(|_| unsupported_rule_graph())?
        .ast_node_type();
    let [cell, number, operator] = formula.ast_node_array.ast_node.as_slice() else {
        return Err(unsupported_rule_graph());
    };
    if cell.ast_node_type != AstNodeType::LinkedCellRefNode as i32
        || number.ast_node_type != AstNodeType::NumberNode as i32
        || number.ast_number_node_number != Some(value)
        || operator.ast_node_type != comparison as i32
    {
        return Err(unsupported_rule_graph());
    }
    Ok(())
}

fn decode_cell_style(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
) -> Result<Option<RgbaColor>> {
    let style = with_object(package, locations, identifier, "cell style", |object| {
        object
            .messages
            .iter()
            .find_map(|message| {
                (message.type_ == CELL_STYLE_MESSAGE_TYPE)
                    .then(|| tst::CellStyleArchive::decode(message.data.as_slice()))
            })
            .transpose()?
            .ok_or_else(|| unsupported_style(identifier))
    })?;
    let Some(mut properties) = style.cell_properties else {
        return Ok(None);
    };
    let fill = properties.cell_fill.take();
    if properties != tst::CellStylePropertiesArchive::default() {
        return Err(unsupported_style(identifier));
    }
    let Some(fill) = fill else {
        return Ok(None);
    };
    if fill.gradient.is_some() || fill.image.is_some() {
        return Err(unsupported_style(identifier));
    }
    fill.color
        .as_ref()
        .map(crate::shapes::color_from_native)
        .transpose()
}

fn decode_text_style(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
) -> Result<(Option<RgbaColor>, bool)> {
    let style = with_object(package, locations, identifier, "text style", |object| {
        object
            .messages
            .iter()
            .find_map(|message| {
                (message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE)
                    .then(|| tswp::ParagraphStyleArchive::decode(message.data.as_slice()))
            })
            .transpose()?
            .ok_or_else(|| unsupported_style(identifier))
    })?;
    if style
        .para_properties
        .is_some_and(|properties| properties != tswp::ParagraphStylePropertiesArchive::default())
    {
        return Err(unsupported_style(identifier));
    }
    let Some(mut properties) = style.char_properties else {
        return Ok((None, false));
    };
    let bold = properties.bold.take().unwrap_or(false);
    let color = properties
        .font_color
        .take()
        .as_ref()
        .map(crate::shapes::color_from_native)
        .transpose()?;
    if properties != tswp::CharacterStylePropertiesArchive::default() {
        return Err(unsupported_style(identifier));
    }
    Ok((color, bold))
}

fn with_object<T>(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
    kind: &str,
    decode: impl FnOnce(&ArchiveObject) -> Result<T>,
) -> Result<T> {
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork conditional-highlight {kind} {identifier} is missing"
        ))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork conditional-highlight {kind} {identifier} is missing"
        ))
    })?;
    decode(object)
}

fn unsupported_rule_graph() -> Error {
    Error::InvalidFormat(
        "iWork conditional-highlight rule uses an unsupported predicate representation".to_owned(),
    )
}

fn unsupported_style(identifier: u64) -> Error {
    Error::InvalidFormat(format!(
        "iWork conditional-highlight style {identifier} uses unsupported overrides"
    ))
}
