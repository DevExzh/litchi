//! Canonical native object-graph encoding for conditional-highlight rules.

use litchi_iwa_core::{ArchiveObject, FieldInfo, FieldPath, FieldType, UnknownFieldRule};

use super::proto::Encode as _;
use super::*;

const CONDITIONAL_STYLE_SET_MESSAGE_TYPE: u32 = 6_010;
const CELL_STYLE_MESSAGE_TYPE: u32 = 6_004;
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const NATIVE_MESSAGE_VERSION: &[u32] = &[1, 0, 5];

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct RuleStyleIdentifiers {
    pub(super) text_style: u64,
    pub(super) cell_style: u64,
}

pub(super) fn conditional_style_graph_objects(
    style_set_id: u64,
    rules: &[Rule],
    formula_owner_uuid: &tsp::Uuid,
    identifiers: &[RuleStyleIdentifiers],
    limits: litchi_iwa_core::Limits,
) -> Result<Vec<ArchiveObject>> {
    if identifiers.len() != rules.len() {
        return Err(invalid_format(
            "conditional-highlight rules and style identifiers differ in length".to_owned(),
        ));
    }
    let mut prepivot = Vec::with_capacity(rules.len());
    let mut current = Vec::with_capacity(rules.len());
    let mut objects = Vec::with_capacity(1 + rules.len() * 2);
    let mut references = Vec::with_capacity(rules.len() * 2);
    for (rule, identifiers) in rules.iter().zip(identifiers) {
        let text_style_id = identifiers.text_style;
        let cell_style_id = identifiers.cell_style;
        let kind = NativePredicateKind::from_condition(&rule.condition);
        let formula = formula::encode(&rule.condition, formula_owner_uuid)?;
        let prepivot_formula = formula::encode_prepivot(&rule.condition, formula_owner_uuid)?;
        let predicate_type = kind.native_value();
        let prepivot_predicate_type = kind.prepivot_native_value();
        let (cell_index, first_index, second_index) = match kind {
            NativePredicateKind::Cell(_) | NativePredicateKind::RelativeDate(_) => (
                PREDICATE_CELL_ARGUMENT_INDEX,
                PREDICATE_UNUSED_ARGUMENT_INDEX,
                PREDICATE_UNUSED_ARGUMENT_INDEX,
            ),
            NativePredicateKind::DatePeriod(kind) => (
                PREDICATE_CELL_ARGUMENT_INDEX,
                formula::date_period_quantity_node_index(kind, &rule.condition)?,
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
        let (first_argument, second_argument) = predicate_arguments(&rule.condition)?;
        let cell_style = tsp::Reference {
            identifier: cell_style_id,
        };
        let text_style = tsp::Reference {
            identifier: text_style_id,
        };
        prepivot.push(
            tst::conditional_style_set_archive::ConditionalStyleRulePrePivot {
                predicate: tst::FormulaPredicatePrePivotArchive {
                    formula: prepivot_formula,
                    predicate_type: prepivot_predicate_type,
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
                        is_spanning_column: Some(false),
                        is_spanning_row: Some(false),
                    }),
                    preserve_row: Some(false),
                    preserve_column: Some(false),
                    ..Default::default()
                }),
                param_value1: Some(first_argument),
                param_value2: Some(second_argument),
                formula: Some(formula),
                for_conditional_style: Some(true),
            }),
            cell_style,
            text_style,
        });
        objects.push(paragraph_style_object(text_style_id, rule, limits)?);
        objects.push(cell_style_object(cell_style_id, rule, limits)?);
        references.push(cell_style_id);
        references.push(text_style_id);
    }
    let style_set = tst::ConditionalStyleSetArchive {
        rule_count: u32::try_from(rules.len()).map_err(|_| {
            invalid_format("conditional-highlight rule count exceeds u32".to_owned())
        })?,
        rules_prepivot: prepivot,
        rules: Some(tst::conditional_style_set_archive::ConditionalStyleRules { rule: current }),
    };
    let mut set_object = structure::new_object(
        style_set_id,
        None,
        CONDITIONAL_STYLE_SET_MESSAGE_TYPE,
        style_set.encode_to_vec(),
        references.clone(),
        limits,
    )?;
    let info = &mut set_object.archive_info.message_infos[0];
    info.versions = NATIVE_MESSAGE_VERSION.to_vec().into();
    let mut field = FieldInfo::new(
        FieldPath::new(vec![3], limits).map_err(|_| Error::InvalidSource)?,
        limits,
    )
    .map_err(|_| Error::InvalidSource)?;
    field.r#type = Some(FieldType::Message);
    field.unknown_field_rule = Some(UnknownFieldRule::IgnoreAndPreserveUntilModified);
    field.object_references = references;
    info.field_infos.push(field);
    objects.insert(0, set_object);
    Ok(objects)
}

pub(super) fn sequential_rule_style_identifiers(
    style_set_id: u64,
    rule_count: usize,
) -> Result<Vec<RuleStyleIdentifiers>> {
    (0..rule_count)
        .map(|index| {
            let offset = u64::try_from(index)
                .ok()
                .and_then(|index| index.checked_mul(2))
                .ok_or(Error::InvalidSource)?;
            let text_style = style_set_id
                .checked_add(offset)
                .and_then(|identifier| identifier.checked_add(1))
                .ok_or(Error::InvalidSource)?;
            let cell_style = text_style.checked_add(1).ok_or(Error::InvalidSource)?;
            Ok(RuleStyleIdentifiers {
                text_style,
                cell_style,
            })
        })
        .collect()
}

fn paragraph_style_object(
    identifier: u64,
    rule: &Rule,
    limits: litchi_iwa_core::Limits,
) -> Result<ArchiveObject> {
    let style = rule.style;
    let text_override_count = u32::from(style.bold()) + u32::from(style.text_color().is_some());
    let data = tswp::ParagraphStyleArchive {
        override_count: (text_override_count != 0).then_some(text_override_count),
        char_properties: (text_override_count != 0).then(|| {
            tswp::CharacterStylePropertiesArchive {
                bold: style.bold().then_some(true),
                font_color: style.text_color().map(color_to_native),
                other: false,
            }
        }),
        para_properties_set: false,
    }
    .encode_to_vec();
    style_object(identifier, PARAGRAPH_STYLE_MESSAGE_TYPE, data, limits)
}

fn cell_style_object(
    identifier: u64,
    rule: &Rule,
    limits: litchi_iwa_core::Limits,
) -> Result<ArchiveObject> {
    let fill = rule.style.fill().map(|color| tsd::FillArchive {
        color: Some(color_to_native(color)),
        other: false,
    });
    let data = tst::CellStyleArchive {
        override_count: fill.is_some().then_some(1),
        cell_properties: fill.map(|cell_fill| tst::CellStylePropertiesArchive {
            cell_fill: Some(cell_fill),
            other: false,
        }),
    }
    .encode_to_vec();
    style_object(identifier, CELL_STYLE_MESSAGE_TYPE, data, limits)
}

fn style_object(
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
    limits: litchi_iwa_core::Limits,
) -> Result<ArchiveObject> {
    let mut object =
        structure::new_object(identifier, None, message_type, data, Vec::new(), limits)?;
    object.archive_info.message_infos[0].versions = NATIVE_MESSAGE_VERSION.to_vec().into();
    Ok(object)
}

fn predicate_arguments(
    condition: &Condition,
) -> Result<(tst::FormulaPredArgArchive, tst::FormulaPredArgArchive)> {
    let none = || tst::FormulaPredArgArchive {
        arg_type: PREDICATE_ARGUMENT_NONE,
        ..Default::default()
    };
    if matches!(
        condition,
        Condition::CellIsBlank
            | Condition::CellIsNotBlank
            | Condition::CheckboxIsChecked
            | Condition::CheckboxIsNotChecked
            | Condition::BooleanIsTrue
            | Condition::BooleanIsFalse
            | Condition::NumberIsPositive
            | Condition::NumberIsNegative
            | Condition::DateIsToday
            | Condition::DateIsYesterday
            | Condition::DateIsTomorrow
    ) {
        return Ok((none(), none()));
    }
    if let Some(value) = condition.date() {
        return Ok((date_argument(value), none()));
    }
    if let Some(range) = condition.date_range() {
        return Ok((date_argument(range.lower()), date_argument(range.upper())));
    }
    if let Some(period) = condition.date_period() {
        return Ok((number_argument(f64::from(period.count()))?, none()));
    }
    if let Some(offset) = condition.date_offset() {
        return Ok((number_argument(f64::from(offset.period().count()))?, none()));
    }
    if let Some(value) = condition.single_operand() {
        return Ok((number_argument(value.get())?, none()));
    }
    if let Some(range) = condition.range() {
        return Ok((
            number_argument(range.lower().get())?,
            number_argument(range.upper().get())?,
        ));
    }
    let text = condition
        .text()
        .expect("text predicate has a text operand")
        .as_str();
    Ok((
        tst::FormulaPredArgArchive {
            arg_type: PREDICATE_ARGUMENT_STRING,
            arg_value: Some(tst::FormulaPredArgDataArchive {
                string_value: Some(text.to_owned()),
                ..Default::default()
            }),
            ..Default::default()
        },
        none(),
    ))
}

fn number_argument(value: f64) -> Result<tst::FormulaPredArgArchive> {
    Ok(tst::FormulaPredArgArchive {
        arg_type: PREDICATE_ARGUMENT_NUMBER,
        arg_value: Some(predicate_number(value)?),
        ..Default::default()
    })
}

fn date_argument(value: Date) -> tst::FormulaPredArgArchive {
    tst::FormulaPredArgArchive {
        arg_type: PREDICATE_ARGUMENT_DATE,
        arg_value: Some(tst::FormulaPredArgDataArchive {
            date_value: Some(value.apple_seconds()),
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn predicate_number(value: f64) -> Result<tst::FormulaPredArgDataArchive> {
    let decimal = litchi_numbers_wire::decimal128_le(value).map_err(|_| Error::InvalidSource)?;
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
