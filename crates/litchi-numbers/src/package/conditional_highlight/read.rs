//! Strict semantic decoding of native conditional-highlight rule graphs.

use super::native::DatePeriodPredicateKind;
use super::proto::Decode as _;
use super::*;
use litchi_iwa_common::table::cell::conditional_highlight::{
    Number, Offset, OffsetDirection, Period, PeriodUnit, Range, Rule, Style, Text,
};

const CONDITIONAL_STYLE_SET_MESSAGE_TYPE: u32 = 6_010;
const CELL_STYLE_MESSAGE_TYPE: u32 = 6_004;
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;

pub(super) fn decode_style_set(
    working: &Working<'_>,
    style_set_id: u64,
) -> Result<tst::ConditionalStyleSetArchive> {
    with_object(working, style_set_id, |object| {
        object
            .messages
            .iter()
            .find(|message| message.type_ == CONDITIONAL_STYLE_SET_MESSAGE_TYPE)
            .map(|message| tst::ConditionalStyleSetArchive::decode(message.data.as_slice()))
            .transpose()?
            .ok_or(Error::InvalidSource)
    })
}

pub(super) fn decode_rules(
    working: &Working<'_>,
    set: &tst::ConditionalStyleSetArchive,
) -> Result<Vec<Rule>> {
    let current = set.rules.as_ref().ok_or_else(unsupported_rule_graph)?;
    let declared = usize::try_from(set.rule_count)
        .map_err(|_| invalid_format("conditional-highlight rule count overflow".to_owned()))?;
    if current.rule.len() != declared
        || (!set.rules_prepivot.is_empty() && set.rules_prepivot.len() != declared)
    {
        return Err(invalid_format(
            "conditional-highlight rule representations disagree on their count".to_owned(),
        ));
    }
    current
        .rule
        .iter()
        .enumerate()
        .map(|(index, rule)| decode_rule(working, rule, set.rules_prepivot.get(index)))
        .collect()
}

fn decode_rule(
    working: &Working<'_>,
    rule: &tst::conditional_style_set_archive::ConditionalStyleRule,
    prepivot: Option<&tst::conditional_style_set_archive::ConditionalStyleRulePrePivot>,
) -> Result<Rule> {
    if prepivot.is_some_and(|prepivot| {
        rule.cell_style.identifier != prepivot.cell_style.identifier
            || rule.text_style.identifier != prepivot.text_style.identifier
    }) {
        return Err(unsupported_rule_graph());
    }
    let predicate = rule.predicate.as_ref().ok_or_else(unsupported_rule_graph)?;
    let condition = decode_predicate(predicate, prepivot.map(|prepivot| &prepivot.predicate))?;
    let fill = decode_cell_style(working, rule.cell_style.identifier)?;
    let (text_color, bold) = decode_text_style(working, rule.text_style.identifier)?;
    Ok(Rule::new(
        condition,
        Style::new(fill, text_color, bold).map_err(|_| {
            invalid_format("conditional-highlight rule has no supported visual override".to_owned())
        })?,
    ))
}

fn decode_predicate(
    predicate: &tst::FormulaPredicateArchive,
    prepivot: Option<&tst::FormulaPredicatePrePivotArchive>,
) -> Result<Condition> {
    let kind = NativePredicateKind::try_from(predicate.predicate_type)
        .map_err(|_| unsupported_rule_graph())?;
    let base_indexes = match kind {
        NativePredicateKind::Cell(_) | NativePredicateKind::RelativeDate(_) => (
            PREDICATE_CELL_ARGUMENT_INDEX,
            PREDICATE_UNUSED_ARGUMENT_INDEX,
            PREDICATE_UNUSED_ARGUMENT_INDEX,
        ),
        NativePredicateKind::DatePeriod(_) => (
            PREDICATE_CELL_ARGUMENT_INDEX,
            PREDICATE_UNUSED_ARGUMENT_INDEX,
            PREDICATE_UNUSED_ARGUMENT_INDEX,
        ),
        NativePredicateKind::Checkbox(_) => (
            PREDICATE_CELL_ARGUMENT_INDEX,
            PREDICATE_NUMBER_ARGUMENT_INDEX,
            PREDICATE_UNUSED_ARGUMENT_INDEX,
        ),
        NativePredicateKind::Boolean(_) => (
            PREDICATE_CELL_ARGUMENT_INDEX,
            PREDICATE_NUMBER_ARGUMENT_INDEX,
            PREDICATE_UNUSED_ARGUMENT_INDEX,
        ),
        NativePredicateKind::FixedDate(FixedDatePredicateKind::Equal) => (
            PREDICATE_DATE_EQUALITY_CELL_ARGUMENT_INDEX,
            PREDICATE_DATE_EQUALITY_ARGUMENT_INDEX,
            PREDICATE_UNUSED_ARGUMENT_INDEX,
        ),
        NativePredicateKind::FixedDate(kind) if kind.is_range() => (
            PREDICATE_RANGE_CELL_ARGUMENT_INDEX,
            PREDICATE_RANGE_LOWER_ARGUMENT_INDEX,
            PREDICATE_RANGE_UPPER_ARGUMENT_INDEX,
        ),
        NativePredicateKind::FixedDate(_) => (
            PREDICATE_CELL_ARGUMENT_INDEX,
            PREDICATE_DATE_ARGUMENT_INDEX,
            PREDICATE_UNUSED_ARGUMENT_INDEX,
        ),
        NativePredicateKind::NumericSign(_) => (
            PREDICATE_CELL_ARGUMENT_INDEX,
            PREDICATE_NUMBER_ARGUMENT_INDEX,
            PREDICATE_UNUSED_ARGUMENT_INDEX,
        ),
        NativePredicateKind::Numeric(kind) if kind.is_range() => (
            PREDICATE_RANGE_CELL_ARGUMENT_INDEX,
            PREDICATE_RANGE_LOWER_ARGUMENT_INDEX,
            PREDICATE_RANGE_UPPER_ARGUMENT_INDEX,
        ),
        NativePredicateKind::Numeric(_) => (
            PREDICATE_CELL_ARGUMENT_INDEX,
            PREDICATE_NUMBER_ARGUMENT_INDEX,
            PREDICATE_UNUSED_ARGUMENT_INDEX,
        ),
        NativePredicateKind::Text(kind) => (
            kind.cell_argument_index(),
            PREDICATE_TEXT_ARGUMENT_INDEX,
            PREDICATE_UNUSED_ARGUMENT_INDEX,
        ),
    };
    if predicate.qualifier1 != PREDICATE_QUALIFIER_NONE
        || predicate.qualifier2 != PREDICATE_QUALIFIER_NONE
        || predicate.for_conditional_style != Some(true)
        || predicate
            .param_value0
            .as_ref()
            .map(|argument| argument.arg_type)
            != Some(PREDICATE_ARGUMENT_RELATIVE_CELL)
    {
        return Err(unsupported_rule_graph());
    }
    if prepivot.is_some_and(|prepivot| {
        kind.prepivot_native_value() != prepivot.predicate_type
            || prepivot.qualifier1 != PREDICATE_QUALIFIER_NONE
            || prepivot.qualifier2 != PREDICATE_QUALIFIER_NONE
    }) {
        return Err(unsupported_rule_graph());
    }
    let condition = match kind {
        NativePredicateKind::Cell(kind) => decode_cell_predicate(predicate, kind)?,
        NativePredicateKind::Checkbox(kind) => decode_checkbox_predicate(predicate, kind)?,
        NativePredicateKind::Boolean(kind) => decode_boolean_predicate(predicate, kind)?,
        NativePredicateKind::FixedDate(kind) => decode_fixed_date_predicate(predicate, kind)?,
        NativePredicateKind::DatePeriod(kind) => decode_date_period_predicate(predicate, kind)?,
        NativePredicateKind::Numeric(kind) => decode_numeric_predicate(predicate, kind)?,
        NativePredicateKind::NumericSign(kind) => decode_sign_predicate(predicate, kind)?,
        NativePredicateKind::RelativeDate(kind) => decode_date_predicate(predicate, kind)?,
        NativePredicateKind::Text(kind) => decode_text_predicate(predicate, kind)?,
    };
    let expected_indexes = match kind {
        NativePredicateKind::DatePeriod(kind) => (
            PREDICATE_CELL_ARGUMENT_INDEX,
            formula::date_period_quantity_node_index(kind, &condition)
                .map_err(|_| unsupported_rule_graph())?,
            PREDICATE_UNUSED_ARGUMENT_INDEX,
        ),
        _ => base_indexes,
    };
    if prepivot.is_some_and(|prepivot| {
        prepivot.param_index0 != expected_indexes.0
            || prepivot.param_index1 != expected_indexes.1
            || prepivot.param_index2 != expected_indexes.2
    }) {
        return Err(unsupported_rule_graph());
    }
    formula::validate(
        predicate
            .formula
            .as_ref()
            .ok_or_else(unsupported_rule_graph)?,
        kind,
        &condition,
    )
    .map_err(|_| unsupported_rule_graph())?;
    if let Some(prepivot) = prepivot {
        formula::validate_prepivot(&prepivot.formula, kind, &condition)
            .map_err(|_| unsupported_rule_graph())?;
    }
    Ok(condition)
}

fn decode_date_period_predicate(
    predicate: &tst::FormulaPredicateArchive,
    kind: DatePeriodPredicateKind,
) -> Result<Condition> {
    if predicate
        .param_value2
        .as_ref()
        .map(|argument| argument.arg_type)
        != Some(PREDICATE_ARGUMENT_NONE)
    {
        return Err(unsupported_rule_graph());
    }
    let value = decode_number_argument(predicate.param_value1.as_ref())?.get();
    if value.fract() != 0.0 || !(1.0..=f64::from(u32::MAX)).contains(&value) {
        return Err(unsupported_rule_graph());
    }
    let count = value as u32;
    let formula = predicate
        .formula
        .as_ref()
        .ok_or_else(unsupported_rule_graph)?;
    let units = [
        PeriodUnit::Days,
        PeriodUnit::Weeks,
        PeriodUnit::Months,
        PeriodUnit::Quarters,
        PeriodUnit::Years,
    ];
    for unit in units {
        let period = Period::new(count, unit).map_err(|_| unsupported_rule_graph())?;
        if let Some(condition) = kind.period_condition(period)
            && formula::validate(formula, NativePredicateKind::DatePeriod(kind), &condition).is_ok()
        {
            return Ok(condition);
        }
        for direction in [OffsetDirection::Ago, OffsetDirection::FromNow] {
            let offset = Offset::new(period, direction);
            if let Some(condition) = kind.offset_condition(offset)
                && formula::validate(formula, NativePredicateKind::DatePeriod(kind), &condition)
                    .is_ok()
            {
                return Ok(condition);
            }
        }
    }
    Err(unsupported_rule_graph())
}

fn decode_checkbox_predicate(
    predicate: &tst::FormulaPredicateArchive,
    kind: CheckboxPredicateKind,
) -> Result<Condition> {
    validate_operand_free_predicate(predicate)?;
    Ok(kind.condition())
}

fn decode_boolean_predicate(
    predicate: &tst::FormulaPredicateArchive,
    kind: BooleanPredicateKind,
) -> Result<Condition> {
    validate_operand_free_predicate(predicate)?;
    Ok(kind.condition())
}

fn decode_cell_predicate(
    predicate: &tst::FormulaPredicateArchive,
    kind: CellPredicateKind,
) -> Result<Condition> {
    validate_operand_free_predicate(predicate)?;
    Ok(kind.condition())
}

fn decode_sign_predicate(
    predicate: &tst::FormulaPredicateArchive,
    kind: NumericSignPredicateKind,
) -> Result<Condition> {
    validate_operand_free_predicate(predicate)?;
    Ok(kind.condition())
}

fn decode_date_predicate(
    predicate: &tst::FormulaPredicateArchive,
    kind: RelativeDatePredicateKind,
) -> Result<Condition> {
    validate_operand_free_predicate(predicate)?;
    Ok(kind.condition())
}

fn decode_fixed_date_predicate(
    predicate: &tst::FormulaPredicateArchive,
    kind: FixedDatePredicateKind,
) -> Result<Condition> {
    let lower = decode_date_argument(predicate.param_value1.as_ref())?;
    if kind.is_range() {
        let upper = decode_date_argument(predicate.param_value2.as_ref())?;
        let range = DateRange::new(lower, upper).map_err(|_| unsupported_rule_graph())?;
        return kind
            .range_condition(range)
            .ok_or_else(unsupported_rule_graph);
    }
    if predicate
        .param_value2
        .as_ref()
        .map(|argument| argument.arg_type)
        != Some(PREDICATE_ARGUMENT_NONE)
    {
        return Err(unsupported_rule_graph());
    }
    kind.condition(lower).ok_or_else(unsupported_rule_graph)
}

fn validate_operand_free_predicate(predicate: &tst::FormulaPredicateArchive) -> Result<()> {
    if [
        predicate.param_value1.as_ref(),
        predicate.param_value2.as_ref(),
    ]
    .into_iter()
    .any(|argument| argument.map(|argument| argument.arg_type) != Some(PREDICATE_ARGUMENT_NONE))
    {
        return Err(unsupported_rule_graph());
    }
    Ok(())
}

fn decode_numeric_predicate(
    predicate: &tst::FormulaPredicateArchive,
    kind: NumericPredicateKind,
) -> Result<Condition> {
    let first = decode_number_argument(predicate.param_value1.as_ref())?;
    let second = if kind.is_range() {
        Some(decode_number_argument(predicate.param_value2.as_ref())?)
    } else {
        if predicate
            .param_value2
            .as_ref()
            .map(|argument| argument.arg_type)
            != Some(PREDICATE_ARGUMENT_NONE)
        {
            return Err(unsupported_rule_graph());
        }
        None
    };
    match (kind, second) {
        (NumericPredicateKind::Between, Some(upper)) => Ok(Condition::Between(
            Range::new(first, upper).map_err(|_| unsupported_rule_graph())?,
        )),
        (NumericPredicateKind::NotBetween, Some(upper)) => Ok(Condition::NotBetween(
            Range::new(first, upper).map_err(|_| unsupported_rule_graph())?,
        )),
        (_, None) => kind
            .single_condition(first)
            .ok_or_else(unsupported_rule_graph),
        _ => Err(unsupported_rule_graph()),
    }
}

fn decode_text_predicate(
    predicate: &tst::FormulaPredicateArchive,
    kind: TextPredicateKind,
) -> Result<Condition> {
    if predicate
        .param_value2
        .as_ref()
        .map(|argument| argument.arg_type)
        != Some(PREDICATE_ARGUMENT_NONE)
    {
        return Err(unsupported_rule_graph());
    }
    let value = predicate
        .param_value1
        .as_ref()
        .filter(|argument| argument.arg_type == PREDICATE_ARGUMENT_STRING)
        .and_then(|argument| argument.arg_value.as_ref())
        .and_then(|value| value.string_value.as_deref())
        .ok_or_else(unsupported_rule_graph)?;
    let text = Text::new(value).map_err(|_| unsupported_rule_graph())?;
    Ok(kind.condition(text))
}

fn decode_number_argument(argument: Option<&tst::FormulaPredArgArchive>) -> Result<Number> {
    let value = argument
        .filter(|argument| argument.arg_type == PREDICATE_ARGUMENT_NUMBER)
        .and_then(|argument| argument.arg_value.as_ref())
        .and_then(|value| value.double_value)
        .ok_or_else(unsupported_rule_graph)?;
    Number::new(value).map_err(|_| unsupported_rule_graph())
}

fn decode_date_argument(argument: Option<&tst::FormulaPredArgArchive>) -> Result<Date> {
    let value = argument
        .filter(|argument| argument.arg_type == PREDICATE_ARGUMENT_DATE)
        .and_then(|argument| argument.arg_value.as_ref())
        .and_then(|value| value.date_value)
        .ok_or_else(unsupported_rule_graph)?;
    Date::new(value).map_err(|_| unsupported_rule_graph())
}

fn decode_cell_style(working: &Working<'_>, identifier: u64) -> Result<Option<RgbaColor>> {
    let style = with_object(working, identifier, |object| {
        object
            .messages
            .iter()
            .find(|message| message.type_ == CELL_STYLE_MESSAGE_TYPE)
            .map(|message| tst::CellStyleArchive::decode(message.data.as_slice()))
            .transpose()?
            .ok_or_else(|| unsupported_style(identifier))
    })?;
    let Some(properties) = style.cell_properties else {
        return Ok(None);
    };
    if properties.other {
        return Err(unsupported_style(identifier));
    }
    let Some(fill) = properties.cell_fill else {
        return Ok(None);
    };
    if fill.other {
        return Err(unsupported_style(identifier));
    }
    fill.color.as_ref().map(color_from_native).transpose()
}

fn decode_text_style(working: &Working<'_>, identifier: u64) -> Result<(Option<RgbaColor>, bool)> {
    let style = with_object(working, identifier, |object| {
        object
            .messages
            .iter()
            .find(|message| message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE)
            .map(|message| tswp::ParagraphStyleArchive::decode(message.data.as_slice()))
            .transpose()?
            .ok_or_else(|| unsupported_style(identifier))
    })?;
    if style.para_properties_set {
        return Err(unsupported_style(identifier));
    }
    let Some(properties) = style.char_properties else {
        return Ok((None, false));
    };
    if properties.other {
        return Err(unsupported_style(identifier));
    }
    let bold = properties.bold.unwrap_or(false);
    let color = properties
        .font_color
        .as_ref()
        .map(color_from_native)
        .transpose()?;
    Ok((color, bold))
}

fn with_object<T>(
    working: &Working<'_>,
    identifier: u64,
    decode: impl FnOnce(&litchi_iwa_core::ArchiveObject) -> Result<T>,
) -> Result<T> {
    let object = working
        .object(identifier)
        .map_err(|_| Error::InvalidSource)?;
    decode(object)
}

fn unsupported_rule_graph() -> Error {
    invalid_format(
        "iWork conditional-highlight rule uses an unsupported predicate representation".to_owned(),
    )
}

fn unsupported_style(identifier: u64) -> Error {
    invalid_format(format!(
        "iWork conditional-highlight style {identifier} uses unsupported overrides"
    ))
}
