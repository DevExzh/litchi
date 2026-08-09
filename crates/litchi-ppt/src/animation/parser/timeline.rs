//! Rotation, scale, set, command, condition, and timeline-control records.

use super::behavior::{parse_time_behavior, parse_time_variant_string, parse_time_visual_element};
use super::support::{
    parse_bool1, read_f32, read_i32, read_u32, require_atom, require_container, require_header,
};
use crate::animation::types::{
    TimeAnimateValueType, TimeBehavior, TimeBehaviorProperty, TimeCommandBehavior,
    TimeCommandBehaviorAtom, TimeCommandBehaviorType, TimeCondition, TimeConditionAtom,
    TimeConditionType, TimeIterateData, TimeIterateDirection, TimeIterateIntervalType,
    TimeIterateType, TimeModifier, TimeRotationBehavior, TimeRotationBehaviorAtom,
    TimeRotationDirection, TimeScaleBehavior, TimeScaleBehaviorAtom, TimeSequenceData,
    TimeSequenceNextAction, TimeSequencePreviousAction, TimeSetBehavior, TimeSetBehaviorAtom,
    TimeTriggerEvent, TimeTriggerObject, is_valid_time_set_value, time_set_attribute_value_type,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

/// Parse an exact rotation behavior container.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_time_rotation_behavior(record: &Record) -> Result<TimeRotationBehavior> {
    require_container(
        record,
        RecordType::TimeRotationBehaviorContainer,
        0,
        "TimeRotationBehaviorContainer",
    )?;
    if record.children.len() != 2 {
        return Err(Error::InvalidFormat(
            "TimeRotationBehaviorContainer requires an atom and common behavior".to_string(),
        ));
    }
    let atom = parse_time_rotation_behavior_atom(&record.children[0])?;
    let behavior = parse_time_behavior(&record.children[1])?;
    if !behavior.atom.attribute_names_used
        || !matches!(
            behavior.attribute_names.as_deref(),
            Some([name]) if matches!(name.as_str(), "r" | "ppt_r")
        )
    {
        return Err(Error::InvalidFormat(
            "rotation behavior requires exactly one r or ppt_r attribute".to_string(),
        ));
    }
    validate_basic_behavior_properties(&behavior)?;
    Ok(TimeRotationBehavior { atom, behavior })
}

/// Parse an exact 20-byte `TimeRotationBehaviorAtom` payload.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_time_rotation_behavior_atom(record: &Record) -> Result<TimeRotationBehaviorAtom> {
    require_atom(
        record,
        RecordType::TimeRotationBehavior,
        0,
        20,
        "TimeRotationBehaviorAtom",
    )?;
    let flags = read_u32(&record.data, 0);
    let by_degrees = (flags & 0x01 != 0).then(|| read_f32(&record.data, 4));
    let from_degrees = if flags & 0x02 != 0 {
        Some(read_f32(&record.data, 8))
    } else if read_f32(&record.data, 8) == 0.0 {
        None
    } else {
        return Err(Error::InvalidFormat(
            "rotation from value must be zero when unused".to_string(),
        ));
    };
    let to_degrees = if flags & 0x04 != 0 {
        Some(read_f32(&record.data, 12))
    } else if read_f32(&record.data, 12).to_bits() == 360.0_f32.to_bits() {
        None
    } else {
        return Err(Error::InvalidFormat(
            "rotation to value must be 360 when unused".to_string(),
        ));
    };
    let direction = if flags & 0x08 != 0 {
        Some(match read_u32(&record.data, 16) {
            0 => TimeRotationDirection::Clockwise,
            1 => TimeRotationDirection::CounterClockwise,
            value => {
                return Err(Error::InvalidFormat(format!(
                    "invalid rotation direction {value}"
                )));
            },
        })
    } else if read_u32(&record.data, 16) == 0 {
        None
    } else {
        return Err(Error::InvalidFormat(
            "rotation direction must be zero when unused".to_string(),
        ));
    };
    if from_degrees.is_some() && by_degrees.is_none() && to_degrees.is_none() {
        return Err(Error::InvalidFormat(
            "rotation from value requires a by or to value".to_string(),
        ));
    }
    Ok(TimeRotationBehaviorAtom {
        by_degrees,
        from_degrees,
        to_degrees,
        direction,
    })
}

/// Parse an exact scale behavior container.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_time_scale_behavior(record: &Record) -> Result<TimeScaleBehavior> {
    require_container(
        record,
        RecordType::TimeScaleBehaviorContainer,
        0,
        "TimeScaleBehaviorContainer",
    )?;
    if record.children.len() != 2 {
        return Err(Error::InvalidFormat(
            "TimeScaleBehaviorContainer requires an atom and common behavior".to_string(),
        ));
    }
    let atom = parse_time_scale_behavior_atom(&record.children[0])?;
    let behavior = parse_time_behavior(&record.children[1])?;
    validate_basic_behavior_properties(&behavior)?;
    Ok(TimeScaleBehavior { atom, behavior })
}

/// Parse an exact 32-byte `TimeScaleBehaviorAtom` payload.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_time_scale_behavior_atom(record: &Record) -> Result<TimeScaleBehaviorAtom> {
    require_atom(
        record,
        RecordType::TimeScaleBehavior,
        0,
        32,
        "TimeScaleBehaviorAtom",
    )?;
    let flags = read_u32(&record.data, 0);
    let by_percent =
        (flags & 0x01 != 0).then(|| (read_f32(&record.data, 4), read_f32(&record.data, 8)));
    let from_percent = if flags & 0x02 != 0 {
        Some((read_f32(&record.data, 12), read_f32(&record.data, 16)))
    } else if read_f32(&record.data, 12) == 0.0 && read_f32(&record.data, 16) == 0.0 {
        None
    } else {
        return Err(Error::InvalidFormat(
            "scale from values must be zero when unused".to_string(),
        ));
    };
    let to_percent = if flags & 0x04 != 0 {
        Some((read_f32(&record.data, 20), read_f32(&record.data, 24)))
    } else if read_f32(&record.data, 20).to_bits() == 100.0_f32.to_bits()
        && read_f32(&record.data, 24).to_bits() == 100.0_f32.to_bits()
    {
        None
    } else {
        return Err(Error::InvalidFormat(
            "scale to values must be 100 when unused".to_string(),
        ));
    };
    let zoom_contents = if flags & 0x08 != 0 {
        Some(parse_bool1(
            record.data[28],
            "TimeScaleBehaviorAtom.fZoomContents",
        )?)
    } else if record.data[28] == 1 {
        None
    } else {
        return Err(Error::InvalidFormat(
            "scale zoom-contents value must be true when unused".to_string(),
        ));
    };
    if from_percent.is_some() && by_percent.is_none() && to_percent.is_none() {
        return Err(Error::InvalidFormat(
            "scale from values require by or to values".to_string(),
        ));
    }
    Ok(TimeScaleBehaviorAtom {
        by_percent,
        from_percent,
        to_percent,
        zoom_contents,
    })
}

/// Parse an exact set-property behavior container.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_time_set_behavior(record: &Record) -> Result<TimeSetBehavior> {
    require_container(
        record,
        RecordType::TimeSetBehaviorContainer,
        0,
        "TimeSetBehaviorContainer",
    )?;
    let atom = record
        .children
        .first()
        .ok_or_else(|| Error::InvalidFormat("set behavior has no atom".to_string()))
        .and_then(parse_time_set_behavior_atom)?;
    let mut index = 1;
    let to =
        if record.children.get(index).is_some_and(|child| {
            child.record_type == RecordType::TimeVariant && child.instance == 1
        }) {
            let value = parse_time_variant_string(&record.children[index])?;
            index += 1;
            Some(value)
        } else {
            None
        };
    let behavior = record
        .children
        .get(index)
        .ok_or_else(|| Error::InvalidFormat("set behavior has no target".to_string()))
        .and_then(parse_time_behavior)?;
    index += 1;
    if index != record.children.len() {
        return Err(Error::InvalidFormat(
            "set behavior has invalid child order or extra children".to_string(),
        ));
    }
    let set = TimeSetBehavior { atom, to, behavior };
    validate_set_behavior(&set)?;
    Ok(set)
}

/// Parse an exact 8-byte `TimeSetBehaviorAtom` payload.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_time_set_behavior_atom(record: &Record) -> Result<TimeSetBehaviorAtom> {
    require_atom(
        record,
        RecordType::TimeSetBehavior,
        0,
        8,
        "TimeSetBehaviorAtom",
    )?;
    let flags = read_u32(&record.data, 0);
    let value = read_u32(&record.data, 4);
    let value_type = if flags & 0x02 != 0 {
        Some(match value {
            0 => TimeAnimateValueType::String,
            1 => TimeAnimateValueType::Number,
            2 => TimeAnimateValueType::Color,
            invalid => {
                return Err(Error::InvalidFormat(format!(
                    "invalid set behavior value type {invalid}"
                )));
            },
        })
    } else if value == 1 {
        None
    } else {
        return Err(Error::InvalidFormat(
            "set behavior value type must be numeric when unused".to_string(),
        ));
    };
    Ok(TimeSetBehaviorAtom {
        to_used: flags & 0x01 != 0,
        value_type,
    })
}

fn validate_set_behavior(set: &TimeSetBehavior) -> Result<()> {
    if set.atom.to_used && set.to.is_none() {
        return Err(Error::InvalidFormat(
            "set to-use flag requires a value".to_string(),
        ));
    }
    if !set.behavior.atom.attribute_names_used {
        return Err(Error::InvalidFormat(
            "set behavior requires an explicit attribute name".to_string(),
        ));
    }
    let attribute = match set.behavior.attribute_names.as_deref() {
        Some([attribute]) => attribute.as_str(),
        _ => {
            return Err(Error::InvalidFormat(
                "set behavior requires exactly one attribute name".to_string(),
            ));
        },
    };
    let expected_type = time_set_attribute_value_type(attribute).ok_or_else(|| {
        Error::InvalidFormat(format!("unsupported set behavior attribute {attribute}"))
    })?;
    let actual_type = set.atom.value_type.unwrap_or(TimeAnimateValueType::Number);
    if actual_type != expected_type {
        return Err(Error::InvalidFormat(
            "set behavior value type does not match its attribute".to_string(),
        ));
    }
    if set
        .to
        .as_deref()
        .is_some_and(|value| !is_valid_time_set_value(attribute, value))
    {
        return Err(Error::InvalidFormat(
            "set behavior value is invalid for its attribute".to_string(),
        ));
    }
    validate_basic_behavior_properties(&set.behavior)
}

pub(super) fn validate_basic_behavior_properties(behavior: &TimeBehavior) -> Result<()> {
    if behavior.properties.as_ref().is_some_and(|list| {
        list.properties.iter().any(|property| {
            matches!(
                property,
                TimeBehaviorProperty::MotionPathEditRelative(_)
                    | TimeBehaviorProperty::ColorModel(_)
                    | TimeBehaviorProperty::ColorDirection(_)
                    | TimeBehaviorProperty::PathEditRotationAngle(_)
                    | TimeBehaviorProperty::PathEditRotationX(_)
                    | TimeBehaviorProperty::PathEditRotationY(_)
                    | TimeBehaviorProperty::PointsTypes(_)
            )
        })
    }) {
        return Err(Error::InvalidFormat(
            "behavior contains properties reserved for color or motion behaviors".to_string(),
        ));
    }
    Ok(())
}

/// Parse an exact command behavior container.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_time_command_behavior(record: &Record) -> Result<TimeCommandBehavior> {
    require_container(
        record,
        RecordType::TimeCommandBehaviorContainer,
        0,
        "TimeCommandBehaviorContainer",
    )?;
    let atom = record
        .children
        .first()
        .ok_or_else(|| Error::InvalidFormat("command behavior has no atom".to_string()))
        .and_then(parse_time_command_behavior_atom)?;
    let mut index = 1;
    let command =
        if record.children.get(index).is_some_and(|child| {
            child.record_type == RecordType::TimeVariant && child.instance == 1
        }) {
            let command = parse_time_variant_string(&record.children[index])?;
            index += 1;
            Some(command)
        } else {
            None
        };
    let behavior = record
        .children
        .get(index)
        .ok_or_else(|| Error::InvalidFormat("command behavior has no target".to_string()))
        .and_then(parse_time_behavior)?;
    index += 1;
    if index != record.children.len() {
        return Err(Error::InvalidFormat(
            "command behavior has invalid child order or extra children".to_string(),
        ));
    }
    validate_basic_behavior_properties(&behavior)?;
    if let Some(command_text) = &command {
        validate_time_command(atom.command_type, command_text)?;
    }
    Ok(TimeCommandBehavior {
        atom,
        command,
        behavior,
    })
}

/// Parse an exact 8-byte `TimeCommandBehaviorAtom` payload.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_time_command_behavior_atom(record: &Record) -> Result<TimeCommandBehaviorAtom> {
    require_atom(
        record,
        RecordType::TimeCommandBehavior,
        0,
        8,
        "TimeCommandBehaviorAtom",
    )?;
    let flags = read_u32(&record.data, 0);
    let value = read_u32(&record.data, 4);
    let command_type = if flags & 0x01 != 0 {
        Some(match value {
            0 => TimeCommandBehaviorType::Event,
            1 => TimeCommandBehaviorType::Call,
            2 => TimeCommandBehaviorType::OleVerb,
            invalid => {
                return Err(Error::InvalidFormat(format!(
                    "invalid command behavior type {invalid}"
                )));
            },
        })
    } else if value == 1 {
        None
    } else {
        return Err(Error::InvalidFormat(
            "command behavior type must be Call when unused".to_string(),
        ));
    };
    Ok(TimeCommandBehaviorAtom {
        command_type,
        command_used: flags & 0x02 != 0,
    })
}

fn validate_time_command(
    command_type: Option<TimeCommandBehaviorType>,
    command: &str,
) -> Result<()> {
    let valid = match command_type.unwrap_or(TimeCommandBehaviorType::Call) {
        TimeCommandBehaviorType::Event => command == "onstopaudio",
        TimeCommandBehaviorType::OleVerb => command.parse::<i32>().is_ok(),
        TimeCommandBehaviorType::Call => {
            matches!(
                command,
                "play" | "pause" | "resume" | "stop" | "togglePause"
            ) || command
                .strip_prefix("playFrom(")
                .and_then(|value| value.strip_suffix(')'))
                .and_then(|value| value.parse::<f64>().ok())
                .is_some_and(|value| value.is_finite() && value >= 0.0)
        },
    };
    if !valid {
        return Err(Error::InvalidFormat(
            "invalid command for command behavior type".to_string(),
        ));
    }
    Ok(())
}

/// Parse an exact `TimeIterateDataAtom`.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_time_iterate_data(record: &Record) -> Result<TimeIterateData> {
    require_atom(
        record,
        RecordType::TimeIterateData,
        0,
        20,
        "TimeIterateDataAtom",
    )?;
    let flags = read_u32(&record.data, 16);
    let interval = optional_u32(flags & 4 != 0, read_u32(&record.data, 0), 0, Ok)?;
    let iterate_type = optional_u32(flags & 2 != 0, read_u32(&record.data, 4), 0, |v| match v {
        0 => Ok(TimeIterateType::AllAtOnce),
        1 => Ok(TimeIterateType::ByWord),
        2 => Ok(TimeIterateType::ByLetter),
        _ => Err(Error::InvalidFormat("invalid iteration type".into())),
    })?;
    let direction = optional_u32(flags & 1 != 0, read_u32(&record.data, 8), 1, |v| match v {
        0 => Ok(TimeIterateDirection::Backward),
        1 => Ok(TimeIterateDirection::Forward),
        _ => Err(Error::InvalidFormat("invalid iteration direction".into())),
    })?;
    let interval_type = optional_u32(flags & 8 != 0, read_u32(&record.data, 12), 0, |v| match v {
        0 => Ok(TimeIterateIntervalType::Milliseconds),
        1 => Ok(TimeIterateIntervalType::TenthsOfAPercent),
        _ => Err(Error::InvalidFormat(
            "invalid iteration interval type".into(),
        )),
    })?;
    Ok(TimeIterateData {
        interval,
        iterate_type,
        direction,
        interval_type,
    })
}

/// Parse an exact `TimeSequenceDataAtom`.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_time_sequence_data(record: &Record) -> Result<TimeSequenceData> {
    require_atom(
        record,
        RecordType::TimeSequenceData,
        0,
        20,
        "TimeSequenceDataAtom",
    )?;
    let flags = read_u32(&record.data, 16);
    let concurrent = optional_u32(flags & 1 != 0, read_u32(&record.data, 0), 0, |v| match v {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(Error::InvalidFormat("invalid sequence concurrency".into())),
    })?;
    let next_action = optional_u32(flags & 2 != 0, read_u32(&record.data, 4), 0, |v| match v {
        0 => Ok(TimeSequenceNextAction::None),
        1 => Ok(TimeSequenceNextAction::SeekToNaturalEnd),
        _ => Err(Error::InvalidFormat("invalid next sequence action".into())),
    })?;
    let previous_action =
        optional_u32(flags & 4 != 0, read_u32(&record.data, 8), 0, |v| match v {
            0 => Ok(TimeSequencePreviousAction::None),
            1 => Ok(TimeSequencePreviousAction::SkipTimedChildren),
            _ => Err(Error::InvalidFormat(
                "invalid previous sequence action".into(),
            )),
        })?;
    Ok(TimeSequenceData {
        concurrent,
        next_action,
        previous_action,
    })
}

fn optional_u32<T>(
    used: bool,
    value: u32,
    default: u32,
    parse: impl FnOnce(u32) -> Result<T>,
) -> Result<Option<T>> {
    if used {
        parse(value).map(Some)
    } else if value == default {
        Ok(None)
    } else {
        Err(Error::InvalidFormat(
            "unused time property has a non-default value".into(),
        ))
    }
}

/// Parse an exact `TimeConditionContainer` and its optional visual target.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_time_condition(record: &Record) -> Result<TimeCondition> {
    require_container(
        record,
        RecordType::TimeConditionContainer,
        record.instance,
        "TimeConditionContainer",
    )?;
    let condition_type = match record.instance {
        1 => TimeConditionType::None,
        2 => TimeConditionType::Begin,
        3 => TimeConditionType::End,
        4 => TimeConditionType::Next,
        5 => TimeConditionType::Previous,
        6 => TimeConditionType::EndSync,
        value => {
            return Err(Error::InvalidFormat(format!(
                "invalid time condition type {value}"
            )));
        },
    };
    let atom = record
        .children
        .first()
        .ok_or_else(|| Error::InvalidFormat("time condition has no atom".to_string()))
        .and_then(parse_time_condition_atom)?;
    let expects_visual = atom.trigger_object == TimeTriggerObject::VisualElement;
    let visual_target = match record.children.get(1) {
        Some(target) if expects_visual => Some(parse_time_visual_element(target)?),
        Some(_) => {
            return Err(Error::InvalidFormat(
                "only visual-element conditions can contain a visual target".to_string(),
            ));
        },
        None if expects_visual => {
            return Err(Error::InvalidFormat(
                "visual-element condition is missing its target".to_string(),
            ));
        },
        None => None,
    };
    if record.children.len() > 2 {
        return Err(Error::InvalidFormat(
            "time condition has extra children".to_string(),
        ));
    }
    Ok(TimeCondition {
        condition_type,
        atom,
        visual_target,
    })
}

/// Parse an exact 16-byte `TimeConditionAtom` payload.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_time_condition_atom(record: &Record) -> Result<TimeConditionAtom> {
    require_atom(
        record,
        RecordType::TimeCondition,
        0,
        16,
        "TimeConditionAtom",
    )?;
    let trigger_object = match read_u32(&record.data, 0) {
        0 => TimeTriggerObject::None,
        1 => TimeTriggerObject::VisualElement,
        2 => TimeTriggerObject::TimeNode,
        3 => TimeTriggerObject::RuntimeNodeReference,
        value => {
            return Err(Error::InvalidFormat(format!(
                "invalid condition trigger object {value}"
            )));
        },
    };
    let trigger_event = match read_u32(&record.data, 4) {
        0 => TimeTriggerEvent::None,
        1 => TimeTriggerEvent::OnBegin,
        3 => TimeTriggerEvent::TimeNodeStart,
        4 => TimeTriggerEvent::TimeNodeEnd,
        5 => TimeTriggerEvent::MouseClick,
        7 => TimeTriggerEvent::MouseOver,
        9 => TimeTriggerEvent::OnNext,
        10 => TimeTriggerEvent::OnPrevious,
        11 => TimeTriggerEvent::StopAudio,
        value => {
            return Err(Error::InvalidFormat(format!(
                "invalid condition trigger event {value}"
            )));
        },
    };
    let target_id = read_u32(&record.data, 8);
    if trigger_object == TimeTriggerObject::RuntimeNodeReference && target_id != 2 {
        return Err(Error::InvalidFormat(
            "runtime-node condition target must be 2".to_string(),
        ));
    }
    Ok(TimeConditionAtom {
        trigger_object,
        trigger_event,
        target_id,
        delay_ms: read_i32(&record.data, 12),
    })
}

/// Parse an exact `TimeModifierAtom`.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_time_modifier(record: &Record) -> Result<TimeModifier> {
    if record.record_type != RecordType::TimeModifier {
        return Err(Error::InvalidFormat(format!(
            "Expected TimeModifierAtom, got {:?}",
            record.record_type
        )));
    }
    require_header(record, 0, record.instance, Some(8), "TimeModifierAtom")?;
    let value = read_u32(&record.data, 4);
    match read_u32(&record.data, 0) {
        0 => Ok(TimeModifier::RepeatCount(value)),
        1 => Ok(TimeModifier::RepeatDuration(value)),
        2 => Ok(TimeModifier::Speed(value)),
        3 => Ok(TimeModifier::Accelerate(value)),
        4 => Ok(TimeModifier::Decelerate(value)),
        5 => Ok(TimeModifier::AutomaticReverse(value)),
        kind => Err(Error::InvalidFormat(format!(
            "invalid time modifier type {kind}"
        ))),
    }
}
