//! Structural parsers for extended and subordinate timing containers.

use super::super::behavior::{
    parse_time_animate_behavior, parse_time_color_behavior, parse_time_effect_behavior,
    parse_time_motion_behavior, parse_time_visual_element,
};
use super::super::support::require_container;
use super::super::timeline::{
    parse_time_command_behavior, parse_time_condition, parse_time_iterate_data,
    parse_time_modifier, parse_time_rotation_behavior, parse_time_scale_behavior,
    parse_time_sequence_data, parse_time_set_behavior,
};
use super::atom::parse_time_node_atom;
use super::model::{NodeParts, SubEffectParts, effective_kind};
use super::properties::parse_time_node_property_list;
use super::validation::{
    ensure_order, set_once, set_subeffect_behavior, set_time_node_behavior, validate_extended_node,
    validate_sub_effect,
};
use crate::animation::types::{
    ExtendedTimeNode, TimeConditionType, TimeNodeBehavior, TimeNodeKind, TimePropertyListContext,
    TimeSubEffect, TimeSubEffectBehavior,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

/// Parse an exact, canonically ordered PowerPoint 2002 extended time node.
pub fn parse_extended_time_node(record: &Record) -> Result<ExtendedTimeNode> {
    require_container(record, RecordType::ExtTimeNode, 1, "ExtTimeNode")?;
    let atom_record = record
        .children
        .first()
        .ok_or_else(|| Error::Corrupted("ExtTimeNode is missing its TimeNodeAtom".to_string()))?;
    let atom = parse_time_node_atom(atom_record)?;
    if record.children[1..]
        .iter()
        .any(|child| child.record_type == RecordType::TimeNode)
    {
        return Err(Error::InvalidFormat(
            "ExtTimeNode contains multiple TimeNodeAtom records".to_string(),
        ));
    }
    let (properties, child_start) = if record
        .children
        .get(1)
        .is_some_and(|child| child.record_type == RecordType::TimePropertyList)
    {
        (
            Some(parse_time_node_property_list(
                &record.children[1],
                TimePropertyListContext::TimeNode,
            )?),
            2,
        )
    } else {
        (None, 1)
    };
    let kind = effective_kind(&atom);
    let mut parts = NodeParts::default();
    let mut last_rank = 1u8;
    for child in &record.children[child_start..] {
        let rank = match child.record_type {
            RecordType::TimeAnimateBehaviorContainer => {
                set_time_node_behavior(
                    &mut parts.behavior,
                    TimeNodeBehavior::Animate(parse_time_animate_behavior(child)?),
                )?;
                2
            },
            RecordType::TimeColorBehaviorContainer => {
                set_time_node_behavior(
                    &mut parts.behavior,
                    TimeNodeBehavior::Color(parse_time_color_behavior(child)?),
                )?;
                3
            },
            RecordType::TimeEffectBehaviorContainer => {
                set_time_node_behavior(
                    &mut parts.behavior,
                    TimeNodeBehavior::Effect(parse_time_effect_behavior(child)?),
                )?;
                4
            },
            RecordType::TimeMotionBehaviorContainer => {
                set_time_node_behavior(
                    &mut parts.behavior,
                    TimeNodeBehavior::Motion(parse_time_motion_behavior(child)?),
                )?;
                5
            },
            RecordType::TimeRotationBehaviorContainer => {
                set_time_node_behavior(
                    &mut parts.behavior,
                    TimeNodeBehavior::Rotation(parse_time_rotation_behavior(child)?),
                )?;
                6
            },
            RecordType::TimeScaleBehaviorContainer => {
                set_time_node_behavior(
                    &mut parts.behavior,
                    TimeNodeBehavior::Scale(parse_time_scale_behavior(child)?),
                )?;
                7
            },
            RecordType::TimeSetBehaviorContainer => {
                set_time_node_behavior(
                    &mut parts.behavior,
                    TimeNodeBehavior::Set(parse_time_set_behavior(child)?),
                )?;
                8
            },
            RecordType::TimeCommandBehaviorContainer => {
                set_time_node_behavior(
                    &mut parts.behavior,
                    TimeNodeBehavior::Command(parse_time_command_behavior(child)?),
                )?;
                9
            },
            RecordType::TimeClientVisualElement => {
                set_once(
                    &mut parts.visual_target,
                    parse_time_visual_element(child)?,
                    "client visual element",
                )?;
                10
            },
            RecordType::TimeIterateData => {
                set_once(
                    &mut parts.iterate_data,
                    parse_time_iterate_data(child)?,
                    "iterate data",
                )?;
                11
            },
            RecordType::TimeSequenceData => {
                set_once(
                    &mut parts.sequence_data,
                    parse_time_sequence_data(child)?,
                    "sequence data",
                )?;
                12
            },
            RecordType::TimeConditionContainer => {
                let condition = parse_time_condition(child)?;
                match condition.condition_type {
                    TimeConditionType::Begin => {
                        parts.begin_conditions.push(condition);
                        13
                    },
                    TimeConditionType::Next if kind == TimeNodeKind::Sequential => {
                        parts.begin_conditions.push(condition);
                        13
                    },
                    TimeConditionType::End => {
                        parts.end_conditions.push(condition);
                        14
                    },
                    TimeConditionType::Previous if kind == TimeNodeKind::Sequential => {
                        parts.end_conditions.push(condition);
                        14
                    },
                    TimeConditionType::EndSync => {
                        set_once(
                            &mut parts.end_sync_condition,
                            condition,
                            "end-sync condition",
                        )?;
                        15
                    },
                    TimeConditionType::Next | TimeConditionType::Previous => {
                        return Err(Error::InvalidFormat(
                            "next/previous conditions require a sequential time node".to_string(),
                        ));
                    },
                    TimeConditionType::None => {
                        return Err(Error::InvalidFormat(
                            "condition type None is not valid in an extended time node".to_string(),
                        ));
                    },
                }
            },
            RecordType::TimeModifier => {
                parts.modifiers.push(parse_time_modifier(child)?);
                16
            },
            RecordType::TimeSubEffectContainer => {
                parts.sub_effects.push(parse_time_sub_effect(child)?);
                17
            },
            RecordType::ExtTimeNode => {
                parts.children.push(parse_extended_time_node(child)?);
                18
            },
            other => {
                return Err(Error::InvalidFormat(format!(
                    "unexpected {other:?} child in ExtTimeNode"
                )));
            },
        };
        ensure_order(&mut last_rank, rank, "ExtTimeNode")?;
    }
    validate_extended_node(kind, &parts)?;
    Ok(parts.finish(atom, properties))
}

/// Parse an exact, canonically ordered subordinate time-node effect.
pub fn parse_time_sub_effect(record: &Record) -> Result<TimeSubEffect> {
    require_container(
        record,
        RecordType::TimeSubEffectContainer,
        1,
        "SubEffectContainer",
    )?;
    let atom = record
        .children
        .first()
        .ok_or_else(|| Error::Corrupted("SubEffectContainer has no TimeNodeAtom".to_string()))
        .and_then(parse_time_node_atom)?;
    let kind = match atom.node_type {
        Some(TimeNodeKind::Behavior) => TimeNodeKind::Behavior,
        Some(TimeNodeKind::Media) => TimeNodeKind::Media,
        _ => {
            return Err(Error::InvalidFormat(
                "subeffect time-node type must explicitly be Behavior or Media".to_string(),
            ));
        },
    };
    let (properties, child_start) = if record
        .children
        .get(1)
        .is_some_and(|child| child.record_type == RecordType::TimePropertyList)
    {
        (
            Some(parse_time_node_property_list(
                &record.children[1],
                TimePropertyListContext::SubEffect,
            )?),
            2,
        )
    } else {
        (None, 1)
    };
    let mut parts = SubEffectParts::default();
    let mut last_rank = 1u8;
    for child in &record.children[child_start..] {
        let rank = match child.record_type {
            RecordType::TimeColorBehaviorContainer => {
                set_subeffect_behavior(
                    &mut parts.behavior,
                    TimeSubEffectBehavior::Color(parse_time_color_behavior(child)?),
                )?;
                2
            },
            RecordType::TimeSetBehaviorContainer => {
                set_subeffect_behavior(
                    &mut parts.behavior,
                    TimeSubEffectBehavior::Set(parse_time_set_behavior(child)?),
                )?;
                3
            },
            RecordType::TimeCommandBehaviorContainer => {
                set_subeffect_behavior(
                    &mut parts.behavior,
                    TimeSubEffectBehavior::Command(parse_time_command_behavior(child)?),
                )?;
                4
            },
            RecordType::TimeClientVisualElement => {
                if parts
                    .visual_target
                    .replace(parse_time_visual_element(child)?)
                    .is_some()
                {
                    return Err(Error::InvalidFormat(
                        "SubEffectContainer has multiple visual targets".to_string(),
                    ));
                }
                5
            },
            RecordType::TimeConditionContainer => {
                let condition = parse_time_condition(child)?;
                match condition.condition_type {
                    TimeConditionType::Begin => {
                        parts.begin_conditions.push(condition);
                        6
                    },
                    TimeConditionType::End => {
                        parts.end_conditions.push(condition);
                        7
                    },
                    _ => {
                        return Err(Error::InvalidFormat(
                            "subeffect conditions must be Begin or End".to_string(),
                        ));
                    },
                }
            },
            RecordType::TimeModifier => {
                parts.modifiers.push(parse_time_modifier(child)?);
                8
            },
            other => {
                return Err(Error::InvalidFormat(format!(
                    "unexpected {other:?} child in SubEffectContainer"
                )));
            },
        };
        ensure_order(&mut last_rank, rank, "SubEffectContainer")?;
    }
    validate_sub_effect(kind, &parts)?;
    Ok(parts.finish(atom, properties))
}
