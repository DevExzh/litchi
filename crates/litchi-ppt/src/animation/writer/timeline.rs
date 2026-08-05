//! Rotation, scale, set, command, and timeline-control records.

use super::behavior::{
    append_time_variant, encode_time_variant_string, write_time_behavior, write_time_visual_element,
};
use super::support::{create_record_header, wrap_record};
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

/// Serialize an exact rotation behavior container.
pub fn write_time_rotation_behavior(behavior: &TimeRotationBehavior) -> Result<Vec<u8>> {
    validate_rotation_behavior(&behavior.behavior)?;
    let mut children = write_time_rotation_behavior_atom(&behavior.atom)?;
    children.extend(write_time_behavior(&behavior.behavior)?);
    wrap_record(RecordType::TimeRotationBehaviorContainer, 0x0F, 0, children)
}

/// Serialize an exact `TimeRotationBehaviorAtom`.
pub fn write_time_rotation_behavior_atom(atom: &TimeRotationBehaviorAtom) -> Result<Vec<u8>> {
    if atom.from_degrees.is_some() && atom.by_degrees.is_none() && atom.to_degrees.is_none() {
        return Err(Error::InvalidFormat(
            "rotation from value requires a by or to value".to_string(),
        ));
    }
    let flags = u32::from(atom.by_degrees.is_some())
        | (u32::from(atom.from_degrees.is_some()) << 1)
        | (u32::from(atom.to_degrees.is_some()) << 2)
        | (u32::from(atom.direction.is_some()) << 3);
    let mut data = Vec::with_capacity(20);
    data.extend(flags.to_le_bytes());
    data.extend(atom.by_degrees.unwrap_or(0.0).to_le_bytes());
    data.extend(atom.from_degrees.unwrap_or(0.0).to_le_bytes());
    data.extend(atom.to_degrees.unwrap_or(360.0).to_le_bytes());
    data.extend(
        atom.direction
            .map_or(0u32, |direction| match direction {
                TimeRotationDirection::Clockwise => 0,
                TimeRotationDirection::CounterClockwise => 1,
            })
            .to_le_bytes(),
    );
    let mut result = create_record_header(RecordType::TimeRotationBehavior, 0, 0, 20);
    result.extend(data);
    Ok(result)
}

/// Serialize an exact scale behavior container.
pub fn write_time_scale_behavior(behavior: &TimeScaleBehavior) -> Result<Vec<u8>> {
    validate_basic_behavior_properties(&behavior.behavior)?;
    let mut children = write_time_scale_behavior_atom(&behavior.atom)?;
    children.extend(write_time_behavior(&behavior.behavior)?);
    wrap_record(RecordType::TimeScaleBehaviorContainer, 0x0F, 0, children)
}

/// Serialize an exact `TimeScaleBehaviorAtom`.
pub fn write_time_scale_behavior_atom(atom: &TimeScaleBehaviorAtom) -> Result<Vec<u8>> {
    if atom.from_percent.is_some() && atom.by_percent.is_none() && atom.to_percent.is_none() {
        return Err(Error::InvalidFormat(
            "scale from values require by or to values".to_string(),
        ));
    }
    let flags = u32::from(atom.by_percent.is_some())
        | (u32::from(atom.from_percent.is_some()) << 1)
        | (u32::from(atom.to_percent.is_some()) << 2)
        | (u32::from(atom.zoom_contents.is_some()) << 3);
    let mut data = Vec::with_capacity(32);
    data.extend(flags.to_le_bytes());
    for value in [
        atom.by_percent.unwrap_or((0.0, 0.0)),
        atom.from_percent.unwrap_or((0.0, 0.0)),
        atom.to_percent.unwrap_or((100.0, 100.0)),
    ] {
        data.extend(value.0.to_le_bytes());
        data.extend(value.1.to_le_bytes());
    }
    data.push(atom.zoom_contents.map_or(1, u8::from));
    data.extend([0, 0, 0]);
    let mut result = create_record_header(RecordType::TimeScaleBehavior, 0, 0, 32);
    result.extend(data);
    Ok(result)
}

/// Serialize an exact set-property behavior container.
pub fn write_time_set_behavior(set: &TimeSetBehavior) -> Result<Vec<u8>> {
    validate_set_behavior(set)?;
    let mut children = write_time_set_behavior_atom(&set.atom);
    if let Some(value) = &set.to {
        append_time_variant(&mut children, 1, encode_time_variant_string(value))?;
    }
    children.extend(write_time_behavior(&set.behavior)?);
    wrap_record(RecordType::TimeSetBehaviorContainer, 0x0F, 0, children)
}

/// Serialize an exact `TimeSetBehaviorAtom`.
pub fn write_time_set_behavior_atom(atom: &TimeSetBehaviorAtom) -> Vec<u8> {
    let flags = u32::from(atom.to_used) | (u32::from(atom.value_type.is_some()) << 1);
    let value_type = atom.value_type.map_or(1u32, |value| match value {
        TimeAnimateValueType::String => 0,
        TimeAnimateValueType::Number => 1,
        TimeAnimateValueType::Color => 2,
    });
    let mut result = create_record_header(RecordType::TimeSetBehavior, 0, 0, 8);
    result.extend(flags.to_le_bytes());
    result.extend(value_type.to_le_bytes());
    result
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

fn validate_rotation_behavior(behavior: &TimeBehavior) -> Result<()> {
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
    validate_basic_behavior_properties(behavior)
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

/// Serialize an exact command behavior container.
pub fn write_time_command_behavior(behavior: &TimeCommandBehavior) -> Result<Vec<u8>> {
    validate_basic_behavior_properties(&behavior.behavior)?;
    if let Some(command) = &behavior.command {
        validate_time_command(behavior.atom.command_type, command)?;
    }
    let mut children = write_time_command_behavior_atom(&behavior.atom);
    if let Some(command) = &behavior.command {
        let data = encode_time_variant_string(command);
        let length = u32::try_from(data.len())
            .map_err(|_| Error::InvalidFormat("time command exceeds 4 GiB".to_string()))?;
        children.extend(create_record_header(RecordType::TimeVariant, 0, 1, length));
        children.extend(data);
    }
    children.extend(write_time_behavior(&behavior.behavior)?);
    wrap_record(RecordType::TimeCommandBehaviorContainer, 0x0F, 0, children)
}

/// Serialize an exact `TimeCommandBehaviorAtom`.
pub fn write_time_command_behavior_atom(atom: &TimeCommandBehaviorAtom) -> Vec<u8> {
    let flags = u32::from(atom.command_type.is_some()) | (u32::from(atom.command_used) << 1);
    let command_type = atom.command_type.map_or(1u32, |value| match value {
        TimeCommandBehaviorType::Event => 0,
        TimeCommandBehaviorType::Call => 1,
        TimeCommandBehaviorType::OleVerb => 2,
    });
    let mut result = create_record_header(RecordType::TimeCommandBehavior, 0, 0, 8);
    result.extend(flags.to_le_bytes());
    result.extend(command_type.to_le_bytes());
    result
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

/// Serialize an exact `TimeIterateDataAtom`.
pub fn write_time_iterate_data(data: &TimeIterateData) -> Vec<u8> {
    let flags = (u32::from(data.direction.is_some()))
        | (u32::from(data.iterate_type.is_some()) << 1)
        | (u32::from(data.interval.is_some()) << 2)
        | (u32::from(data.interval_type.is_some()) << 3);
    let mut payload = Vec::with_capacity(20);
    payload.extend(data.interval.unwrap_or(0).to_le_bytes());
    payload.extend(
        data.iterate_type
            .map_or(0u32, |v| match v {
                TimeIterateType::AllAtOnce => 0,
                TimeIterateType::ByWord => 1,
                TimeIterateType::ByLetter => 2,
            })
            .to_le_bytes(),
    );
    payload.extend(
        data.direction
            .map_or(1u32, |v| match v {
                TimeIterateDirection::Backward => 0,
                TimeIterateDirection::Forward => 1,
            })
            .to_le_bytes(),
    );
    payload.extend(
        data.interval_type
            .map_or(0u32, |v| match v {
                TimeIterateIntervalType::Milliseconds => 0,
                TimeIterateIntervalType::TenthsOfAPercent => 1,
            })
            .to_le_bytes(),
    );
    payload.extend(flags.to_le_bytes());
    let mut result = create_record_header(RecordType::TimeIterateData, 0, 0, 20);
    result.extend(payload);
    result
}

/// Serialize an exact `TimeSequenceDataAtom`.
pub fn write_time_sequence_data(data: &TimeSequenceData) -> Vec<u8> {
    let flags = u32::from(data.concurrent.is_some())
        | (u32::from(data.next_action.is_some()) << 1)
        | (u32::from(data.previous_action.is_some()) << 2);
    let mut payload = Vec::with_capacity(20);
    payload.extend(data.concurrent.map_or(0u32, u32::from).to_le_bytes());
    payload.extend(
        data.next_action
            .map_or(0u32, |v| match v {
                TimeSequenceNextAction::None => 0,
                TimeSequenceNextAction::SeekToNaturalEnd => 1,
            })
            .to_le_bytes(),
    );
    payload.extend(
        data.previous_action
            .map_or(0u32, |v| match v {
                TimeSequencePreviousAction::None => 0,
                TimeSequencePreviousAction::SkipTimedChildren => 1,
            })
            .to_le_bytes(),
    );
    payload.extend(0u32.to_le_bytes());
    payload.extend(flags.to_le_bytes());
    let mut result = create_record_header(RecordType::TimeSequenceData, 0, 0, 20);
    result.extend(payload);
    result
}

/// Serialize an exact `TimeConditionContainer`.
pub fn write_time_condition(condition: &TimeCondition) -> Result<Vec<u8>> {
    let expects_visual = condition.atom.trigger_object == TimeTriggerObject::VisualElement;
    if expects_visual != condition.visual_target.is_some() {
        return Err(Error::InvalidFormat(
            "visual target must exist exactly for visual-element conditions".to_string(),
        ));
    }
    let mut children = write_time_condition_atom(&condition.atom)?;
    if let Some(target) = &condition.visual_target {
        children.extend(write_time_visual_element(target)?);
    }
    let instance = match condition.condition_type {
        TimeConditionType::None => 1,
        TimeConditionType::Begin => 2,
        TimeConditionType::End => 3,
        TimeConditionType::Next => 4,
        TimeConditionType::Previous => 5,
        TimeConditionType::EndSync => 6,
    };
    wrap_record(RecordType::TimeConditionContainer, 0x0F, instance, children)
}

/// Serialize an exact `TimeConditionAtom`.
pub fn write_time_condition_atom(atom: &TimeConditionAtom) -> Result<Vec<u8>> {
    if atom.trigger_object == TimeTriggerObject::RuntimeNodeReference && atom.target_id != 2 {
        return Err(Error::InvalidFormat(
            "runtime-node condition target must be 2".to_string(),
        ));
    }
    let trigger_object = match atom.trigger_object {
        TimeTriggerObject::None => 0u32,
        TimeTriggerObject::VisualElement => 1,
        TimeTriggerObject::TimeNode => 2,
        TimeTriggerObject::RuntimeNodeReference => 3,
    };
    let trigger_event = match atom.trigger_event {
        TimeTriggerEvent::None => 0u32,
        TimeTriggerEvent::OnBegin => 1,
        TimeTriggerEvent::TimeNodeStart => 3,
        TimeTriggerEvent::TimeNodeEnd => 4,
        TimeTriggerEvent::MouseClick => 5,
        TimeTriggerEvent::MouseOver => 7,
        TimeTriggerEvent::OnNext => 9,
        TimeTriggerEvent::OnPrevious => 10,
        TimeTriggerEvent::StopAudio => 11,
    };
    let mut result = create_record_header(RecordType::TimeCondition, 0, 0, 16);
    result.extend(trigger_object.to_le_bytes());
    result.extend(trigger_event.to_le_bytes());
    result.extend(atom.target_id.to_le_bytes());
    result.extend(atom.delay_ms.to_le_bytes());
    Ok(result)
}

/// Serialize an exact `TimeModifierAtom`.
pub fn write_time_modifier(modifier: &TimeModifier) -> Vec<u8> {
    let (kind, value) = match modifier {
        TimeModifier::RepeatCount(value) => (0u32, *value),
        TimeModifier::RepeatDuration(value) => (1, *value),
        TimeModifier::Speed(value) => (2, *value),
        TimeModifier::Accelerate(value) => (3, *value),
        TimeModifier::Decelerate(value) => (4, *value),
        TimeModifier::AutomaticReverse(value) => (5, *value),
    };
    let mut result = create_record_header(RecordType::TimeModifier, 0, 0, 8);
    result.extend(kind.to_le_bytes());
    result.extend(value.to_le_bytes());
    result
}
