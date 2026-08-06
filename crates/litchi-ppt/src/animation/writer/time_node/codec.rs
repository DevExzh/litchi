//! PowerPoint record codecs for timing-node containers and atoms.

use super::super::behavior::{
    write_time_animate_behavior, write_time_color_behavior, write_time_effect_behavior,
    write_time_motion_behavior, write_time_visual_element,
};
use super::super::support::{create_record_header, wrap_record};
use super::super::timeline::{
    write_time_command_behavior, write_time_condition, write_time_iterate_data,
    write_time_modifier, write_time_rotation_behavior, write_time_scale_behavior,
    write_time_sequence_data, write_time_set_behavior,
};
use super::model::{NodeView, SubEffectView};
use super::properties::write_time_node_property_list;
use super::validation::{validate_extended_time_node, validate_sub_effect};
use crate::animation::types::{
    ExtendedTimeNode, TimeNodeAtom, TimeNodeBehavior, TimePropertyListContext, TimeSubEffect,
    TimeSubEffectBehavior,
};
use crate::consts::RecordType;
use crate::package::Result;

/// Serialize a canonically ordered PowerPoint 2002 extended time node.
pub fn write_extended_time_node(node: &ExtendedTimeNode) -> Result<Vec<u8>> {
    let view = NodeView::new(node);
    validate_extended_time_node(&view)?;
    let node = view.source();
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

/// Serialize a canonically ordered subordinate time-node effect.
pub fn write_time_sub_effect(sub_effect: &TimeSubEffect) -> Result<Vec<u8>> {
    let view = SubEffectView::new(sub_effect);
    validate_sub_effect(&view)?;
    let sub_effect = view.source();
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
