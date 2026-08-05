//! Extended time-node envelopes, subordinate effects, and property lists.

use super::behavior::{
    write_time_animate_behavior, write_time_color_behavior, write_time_effect_behavior,
    write_time_motion_behavior, write_time_visual_element,
};
use super::support::{create_record_header, wrap_record};
use super::timeline::{
    write_time_command_behavior, write_time_condition, write_time_iterate_data,
    write_time_modifier, write_time_rotation_behavior, write_time_scale_behavior,
    write_time_sequence_data, write_time_set_behavior,
};
use crate::animation::types::{
    ExtendedTimeNode, TimeConditionType, TimeEffectNodeType, TimeEffectType, TimeMasterRelation,
    TimeNodeAtom, TimeNodeBehavior, TimeNodeKind, TimeNodeProperty, TimeNodePropertyList,
    TimePropertyListContext, TimeSubEffect, TimeSubEffectBehavior,
    has_valid_time_effect_properties, is_valid_time_filter,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};

/// Serialize a canonically ordered PowerPoint 2002 extended time node.
pub fn write_extended_time_node(node: &ExtendedTimeNode) -> Result<Vec<u8>> {
    validate_extended_time_node(node)?;
    let mut children = write_time_node_atom(&node.atom);
    if let Some(properties) = &node.properties {
        children.extend(write_time_node_property_list(
            properties,
            TimePropertyListContext::TimeNode,
        )?);
    }
    if let Some(behavior) = &node.behavior {
        children.extend(match behavior {
            TimeNodeBehavior::Animate(value) => write_time_animate_behavior(value)?,
            TimeNodeBehavior::Color(value) => write_time_color_behavior(value)?,
            TimeNodeBehavior::Effect(value) => write_time_effect_behavior(value)?,
            TimeNodeBehavior::Motion(value) => write_time_motion_behavior(value)?,
            TimeNodeBehavior::Rotation(value) => write_time_rotation_behavior(value)?,
            TimeNodeBehavior::Scale(value) => write_time_scale_behavior(value)?,
            TimeNodeBehavior::Set(value) => write_time_set_behavior(value)?,
            TimeNodeBehavior::Command(value) => write_time_command_behavior(value)?,
        });
    }
    if let Some(target) = &node.visual_target {
        children.extend(write_time_visual_element(target)?);
    }
    if let Some(data) = &node.iterate_data {
        children.extend(write_time_iterate_data(data));
    }
    if let Some(data) = &node.sequence_data {
        children.extend(write_time_sequence_data(data));
    }
    for condition in &node.begin_conditions {
        children.extend(write_time_condition(condition)?);
    }
    for condition in &node.end_conditions {
        children.extend(write_time_condition(condition)?);
    }
    if let Some(condition) = &node.end_sync_condition {
        children.extend(write_time_condition(condition)?);
    }
    for modifier in &node.modifiers {
        children.extend(write_time_modifier(modifier));
    }
    for sub_effect in &node.sub_effects {
        children.extend(write_time_sub_effect(sub_effect)?);
    }
    for child in &node.children {
        children.extend(write_extended_time_node(child)?);
    }
    wrap_record(RecordType::ExtTimeNode, 0x0F, 1, children)
}

fn validate_extended_time_node(node: &ExtendedTimeNode) -> Result<()> {
    let kind = node.atom.node_type.unwrap_or(TimeNodeKind::Parallel);
    if node.behavior.is_some() && kind != TimeNodeKind::Behavior {
        return Err(Error::InvalidFormat(
            "animation behaviors require a behavior time node".to_string(),
        ));
    }
    if node.visual_target.is_some() && kind != TimeNodeKind::Media {
        return Err(Error::InvalidFormat(
            "standalone visual targets require a media time node".to_string(),
        ));
    }
    if node.sequence_data.is_some() && kind != TimeNodeKind::Sequential {
        return Err(Error::InvalidFormat(
            "sequence data requires a sequential time node".to_string(),
        ));
    }
    for condition in &node.begin_conditions {
        let valid = condition.condition_type == TimeConditionType::Begin
            || (kind == TimeNodeKind::Sequential
                && condition.condition_type == TimeConditionType::Next);
        if !valid {
            return Err(Error::InvalidFormat(
                "begin-condition arrays may contain only begin conditions, or next conditions on sequential nodes"
                    .to_string(),
            ));
        }
    }
    for condition in &node.end_conditions {
        let valid = condition.condition_type == TimeConditionType::End
            || (kind == TimeNodeKind::Sequential
                && condition.condition_type == TimeConditionType::Previous);
        if !valid {
            return Err(Error::InvalidFormat(
                "end-condition arrays may contain only end conditions, or previous conditions on sequential nodes"
                    .to_string(),
            ));
        }
    }
    if node
        .end_sync_condition
        .as_ref()
        .is_some_and(|condition| condition.condition_type != TimeConditionType::EndSync)
    {
        return Err(Error::InvalidFormat(
            "end-sync condition must use the EndSync condition type".to_string(),
        ));
    }
    Ok(())
}

/// Serialize a canonically ordered subordinate time-node effect.
pub fn write_time_sub_effect(sub_effect: &TimeSubEffect) -> Result<Vec<u8>> {
    let kind = match sub_effect.atom.node_type {
        Some(TimeNodeKind::Behavior) => TimeNodeKind::Behavior,
        Some(TimeNodeKind::Media) => TimeNodeKind::Media,
        _ => {
            return Err(Error::InvalidFormat(
                "subeffect time-node type must explicitly be Behavior or Media".to_string(),
            ));
        },
    };
    if sub_effect.behavior.is_some() && kind != TimeNodeKind::Behavior {
        return Err(Error::InvalidFormat(
            "subeffect behavior requires a behavior time node".to_string(),
        ));
    }
    if sub_effect.visual_target.is_some() && kind != TimeNodeKind::Media {
        return Err(Error::InvalidFormat(
            "subeffect visual target requires a media time node".to_string(),
        ));
    }
    if sub_effect
        .begin_conditions
        .iter()
        .any(|condition| condition.condition_type != TimeConditionType::Begin)
        || sub_effect
            .end_conditions
            .iter()
            .any(|condition| condition.condition_type != TimeConditionType::End)
    {
        return Err(Error::InvalidFormat(
            "subeffect condition arrays must contain only their matching Begin or End type"
                .to_string(),
        ));
    }

    let mut children = write_time_node_atom(&sub_effect.atom);
    if let Some(properties) = &sub_effect.properties {
        children.extend(write_time_node_property_list(
            properties,
            TimePropertyListContext::SubEffect,
        )?);
    }
    if let Some(behavior) = &sub_effect.behavior {
        children.extend(match behavior {
            TimeSubEffectBehavior::Color(value) => write_time_color_behavior(value)?,
            TimeSubEffectBehavior::Set(value) => write_time_set_behavior(value)?,
            TimeSubEffectBehavior::Command(value) => write_time_command_behavior(value)?,
        });
    }
    if let Some(target) = &sub_effect.visual_target {
        children.extend(write_time_visual_element(target)?);
    }
    for condition in &sub_effect.begin_conditions {
        children.extend(write_time_condition(condition)?);
    }
    for condition in &sub_effect.end_conditions {
        children.extend(write_time_condition(condition)?);
    }
    for modifier in &sub_effect.modifiers {
        children.extend(write_time_modifier(modifier));
    }
    wrap_record(RecordType::TimeSubEffectContainer, 0x0F, 1, children)
}

/// Serialize an exact 32-byte `TimeNodeAtom` payload.
pub fn write_time_node_atom(atom: &TimeNodeAtom) -> Vec<u8> {
    let mut data = Vec::with_capacity(32);
    data.extend(0u32.to_le_bytes());
    data.extend(atom.restart.map_or(0, |value| value.as_u32()).to_le_bytes());
    data.extend(
        atom.node_type
            .map_or(0, |value| value.as_u32())
            .to_le_bytes(),
    );
    data.extend(atom.fill.map_or(0, |value| value.as_u32()).to_le_bytes());
    data.extend(0u32.to_le_bytes());
    data.extend(0u32.to_le_bytes());
    data.extend(atom.duration_ms.unwrap_or(0).to_le_bytes());
    let flags = u32::from(atom.fill.is_some())
        | (u32::from(atom.restart.is_some()) << 1)
        | (u32::from(atom.node_type.is_some()) << 3)
        | (u32::from(atom.duration_ms.is_some()) << 4);
    data.extend(flags.to_le_bytes());
    let mut result = create_record_header(RecordType::TimeNode, 0, 0, 32);
    result.extend(data);
    result
}

/// Serialize a typed `TimePropertyList4TimeNodeContainer`.
pub fn write_time_node_property_list(
    list: &TimeNodePropertyList,
    context: TimePropertyListContext,
) -> Result<Vec<u8>> {
    if !has_valid_time_effect_properties(&list.properties) {
        return Err(Error::InvalidFormat(
            "invalid effect ID, type, or direction combination".to_string(),
        ));
    }
    let mut seen = std::collections::HashSet::with_capacity(list.properties.len());
    let has_interactive_sequence = list.properties.iter().any(|property| {
        matches!(
            property,
            TimeNodeProperty::EffectNodeType(TimeEffectNodeType::InteractiveSequence)
        )
    });
    let mut children = Vec::new();
    for property in &list.properties {
        let (id, data) = encode_time_node_property(property)?;
        if !seen.insert(id) {
            return Err(Error::InvalidFormat(format!(
                "duplicate time property {id:#X}"
            )));
        }
        validate_time_property_context(id, context)?;
        if matches!(property, TimeNodeProperty::EventFilter(_)) && !has_interactive_sequence {
            return Err(Error::InvalidFormat(
                "event filter requires an interactive sequence".to_string(),
            ));
        }
        let length = u32::try_from(data.len()).map_err(|_| {
            Error::InvalidFormat("time property exceeds 4 GiB record limit".to_string())
        })?;
        children.extend(create_record_header(RecordType::TimeVariant, 0, id, length));
        children.extend(data);
    }
    wrap_record(RecordType::TimePropertyList, 0x0F, 0, children)
}

fn encode_time_node_property(property: &TimeNodeProperty) -> Result<(u16, Vec<u8>)> {
    let integer = |value: i32| {
        let mut data = vec![1];
        data.extend(value.to_le_bytes());
        data
    };
    let boolean = |value: bool| vec![0, u8::from(value)];
    let string = |value: &str| {
        let mut data = Vec::with_capacity(1 + value.len().saturating_mul(2));
        data.push(3);
        data.extend(value.encode_utf16().flat_map(u16::to_le_bytes));
        data
    };
    Ok(match property {
        TimeNodeProperty::DisplayHidden(value) => (0x02, integer(i32::from(*value))),
        TimeNodeProperty::MasterRelation(value) => (
            0x05,
            integer(match value {
                TimeMasterRelation::DoNotStart => 0,
                TimeMasterRelation::StartWithMaster => 2,
            }),
        ),
        TimeNodeProperty::SubType => (0x06, integer(1)),
        TimeNodeProperty::EffectId(value) => (0x09, integer(*value)),
        TimeNodeProperty::EffectDirection(value) => (0x0A, integer(*value)),
        TimeNodeProperty::EffectType(value) => (
            0x0B,
            integer(match value {
                TimeEffectType::Entrance => 1,
                TimeEffectType::Exit => 2,
                TimeEffectType::Emphasis => 3,
                TimeEffectType::MotionPath => 4,
                TimeEffectType::ActionVerb => 5,
                TimeEffectType::MediaCommand => 6,
            }),
        ),
        TimeNodeProperty::AfterEffect(value) => (0x0D, boolean(*value)),
        TimeNodeProperty::SlideCount(value) => (0x0F, integer(*value)),
        TimeNodeProperty::TimeFilter(value) => {
            if !is_valid_time_filter(value) {
                return Err(Error::InvalidFormat("invalid time filter".to_string()));
            }
            (0x10, string(value))
        },
        TimeNodeProperty::EventFilter(value) => {
            if value != "cancelBubble" {
                return Err(Error::InvalidFormat(
                    "event filter must be cancelBubble".to_string(),
                ));
            }
            (0x11, string(value))
        },
        TimeNodeProperty::HideWhenStopped(value) => (0x12, boolean(*value)),
        TimeNodeProperty::GroupId(value) => (0x13, integer(*value)),
        TimeNodeProperty::EffectNodeType(value) => (
            0x14,
            integer(match value {
                TimeEffectNodeType::ClickEffect => 1,
                TimeEffectNodeType::WithPrevious => 2,
                TimeEffectNodeType::AfterPrevious => 3,
                TimeEffectNodeType::MainSequence => 4,
                TimeEffectNodeType::InteractiveSequence => 5,
                TimeEffectNodeType::ClickParallel => 6,
                TimeEffectNodeType::WithGroup => 7,
                TimeEffectNodeType::AfterGroup => 8,
                TimeEffectNodeType::TimingRoot => 9,
            }),
        ),
        TimeNodeProperty::PlaceholderNode(value) => (0x15, boolean(*value)),
        TimeNodeProperty::MediaVolume(value) => {
            if !value.is_finite() || !(0.0..=100_000.0).contains(value) {
                return Err(Error::InvalidFormat(
                    "media volume out of range".to_string(),
                ));
            }
            let mut data = vec![2];
            data.extend(value.to_le_bytes());
            (0x16, data)
        },
        TimeNodeProperty::MediaMute(value) => (0x17, boolean(*value)),
        TimeNodeProperty::ZoomToFullScreen(value) => (0x1A, boolean(*value)),
    })
}

fn validate_time_property_context(id: u16, context: TimePropertyListContext) -> Result<()> {
    let invalid = match context {
        TimePropertyListContext::TimeNode => matches!(id, 0x05 | 0x06),
        TimePropertyListContext::SubEffect => {
            matches!(id, 0x09..=0x0B | 0x0F..=0x14 | 0x1A)
        },
    };
    if invalid {
        return Err(Error::InvalidFormat(format!(
            "time property {id:#X} is invalid in {context:?} context"
        )));
    }
    Ok(())
}
