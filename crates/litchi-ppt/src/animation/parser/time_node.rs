//! Extended time-node envelopes, subordinate effects, and slide extensions.

use super::behavior::{
    parse_time_animate_behavior, parse_time_color_behavior, parse_time_effect_behavior,
    parse_time_motion_behavior, parse_time_visual_element, require_time_variant_payload,
};
use super::build::parse_build_list;
use super::support::{
    parse_bool1, parse_optional_time_value, read_u32, require_atom, require_container,
};
use super::timeline::{
    parse_time_command_behavior, parse_time_condition, parse_time_iterate_data,
    parse_time_modifier, parse_time_rotation_behavior, parse_time_scale_behavior,
    parse_time_sequence_data, parse_time_set_behavior,
};
use crate::animation::linked_slide::{LinkedShape, LinkedSlide};
use crate::animation::slide_metadata::SlideTime;
use crate::animation::types::{
    ExtendedTimeNode, Flags, SlideAnimationExtension, TimeConditionType, TimeEffectNodeType,
    TimeEffectType, TimeMasterRelation, TimeNodeAtom, TimeNodeBehavior, TimeNodeFill, TimeNodeKind,
    TimeNodeProperty, TimeNodePropertyList, TimeNodeRestart, TimePropertyListContext,
    TimeSubEffect, TimeSubEffectBehavior, has_valid_time_effect_properties, is_valid_time_filter,
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
    let effective_kind = atom.node_type.unwrap_or(TimeNodeKind::Parallel);
    let mut behavior = None;
    let mut visual_target = None;
    let mut iterate_data = None;
    let mut sequence_data = None;
    let mut begin_conditions = Vec::new();
    let mut end_conditions = Vec::new();
    let mut end_sync_condition = None;
    let mut modifiers = Vec::new();
    let mut sub_effects = Vec::new();
    let mut children = Vec::new();
    let mut last_rank = 1u8;
    for child in &record.children[child_start..] {
        let rank = match child.record_type {
            RecordType::TimeAnimateBehaviorContainer => {
                set_time_node_behavior(
                    &mut behavior,
                    TimeNodeBehavior::Animate(parse_time_animate_behavior(child)?),
                )?;
                2
            },
            RecordType::TimeColorBehaviorContainer => {
                set_time_node_behavior(
                    &mut behavior,
                    TimeNodeBehavior::Color(parse_time_color_behavior(child)?),
                )?;
                3
            },
            RecordType::TimeEffectBehaviorContainer => {
                set_time_node_behavior(
                    &mut behavior,
                    TimeNodeBehavior::Effect(parse_time_effect_behavior(child)?),
                )?;
                4
            },
            RecordType::TimeMotionBehaviorContainer => {
                set_time_node_behavior(
                    &mut behavior,
                    TimeNodeBehavior::Motion(parse_time_motion_behavior(child)?),
                )?;
                5
            },
            RecordType::TimeRotationBehaviorContainer => {
                set_time_node_behavior(
                    &mut behavior,
                    TimeNodeBehavior::Rotation(parse_time_rotation_behavior(child)?),
                )?;
                6
            },
            RecordType::TimeScaleBehaviorContainer => {
                set_time_node_behavior(
                    &mut behavior,
                    TimeNodeBehavior::Scale(parse_time_scale_behavior(child)?),
                )?;
                7
            },
            RecordType::TimeSetBehaviorContainer => {
                set_time_node_behavior(
                    &mut behavior,
                    TimeNodeBehavior::Set(parse_time_set_behavior(child)?),
                )?;
                8
            },
            RecordType::TimeCommandBehaviorContainer => {
                set_time_node_behavior(
                    &mut behavior,
                    TimeNodeBehavior::Command(parse_time_command_behavior(child)?),
                )?;
                9
            },
            RecordType::TimeClientVisualElement => {
                set_once(
                    &mut visual_target,
                    parse_time_visual_element(child)?,
                    "client visual element",
                )?;
                10
            },
            RecordType::TimeIterateData => {
                set_once(
                    &mut iterate_data,
                    parse_time_iterate_data(child)?,
                    "iterate data",
                )?;
                11
            },
            RecordType::TimeSequenceData => {
                set_once(
                    &mut sequence_data,
                    parse_time_sequence_data(child)?,
                    "sequence data",
                )?;
                12
            },
            RecordType::TimeConditionContainer => {
                let condition = parse_time_condition(child)?;
                match condition.condition_type {
                    TimeConditionType::Begin => {
                        begin_conditions.push(condition);
                        13
                    },
                    TimeConditionType::Next if effective_kind == TimeNodeKind::Sequential => {
                        begin_conditions.push(condition);
                        13
                    },
                    TimeConditionType::End => {
                        end_conditions.push(condition);
                        14
                    },
                    TimeConditionType::Previous if effective_kind == TimeNodeKind::Sequential => {
                        end_conditions.push(condition);
                        14
                    },
                    TimeConditionType::EndSync => {
                        set_once(&mut end_sync_condition, condition, "end-sync condition")?;
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
                modifiers.push(parse_time_modifier(child)?);
                16
            },
            RecordType::TimeSubEffectContainer => {
                sub_effects.push(parse_time_sub_effect(child)?);
                17
            },
            RecordType::ExtTimeNode => {
                children.push(parse_extended_time_node(child)?);
                18
            },
            other => {
                return Err(Error::InvalidFormat(format!(
                    "unexpected {other:?} child in ExtTimeNode"
                )));
            },
        };
        if rank < last_rank {
            return Err(Error::InvalidFormat(
                "ExtTimeNode children are not in canonical order".to_string(),
            ));
        }
        last_rank = rank;
    }
    if behavior.is_some() && effective_kind != TimeNodeKind::Behavior {
        return Err(Error::InvalidFormat(
            "animation behaviors require a behavior time node".to_string(),
        ));
    }
    if visual_target.is_some() && effective_kind != TimeNodeKind::Media {
        return Err(Error::InvalidFormat(
            "standalone visual targets require a media time node".to_string(),
        ));
    }
    if sequence_data.is_some() && effective_kind != TimeNodeKind::Sequential {
        return Err(Error::InvalidFormat(
            "sequence data requires a sequential time node".to_string(),
        ));
    }
    Ok(ExtendedTimeNode {
        atom,
        properties,
        behavior,
        visual_target,
        iterate_data,
        sequence_data,
        begin_conditions,
        end_conditions,
        end_sync_condition,
        modifiers,
        sub_effects,
        children,
    })
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
    let mut behavior = None;
    let mut visual_target = None;
    let mut begin_conditions = Vec::new();
    let mut end_conditions = Vec::new();
    let mut modifiers = Vec::new();
    let mut last_rank = 1u8;
    for child in &record.children[child_start..] {
        let rank = match child.record_type {
            RecordType::TimeColorBehaviorContainer => {
                set_subeffect_behavior(
                    &mut behavior,
                    TimeSubEffectBehavior::Color(parse_time_color_behavior(child)?),
                )?;
                2
            },
            RecordType::TimeSetBehaviorContainer => {
                set_subeffect_behavior(
                    &mut behavior,
                    TimeSubEffectBehavior::Set(parse_time_set_behavior(child)?),
                )?;
                3
            },
            RecordType::TimeCommandBehaviorContainer => {
                set_subeffect_behavior(
                    &mut behavior,
                    TimeSubEffectBehavior::Command(parse_time_command_behavior(child)?),
                )?;
                4
            },
            RecordType::TimeClientVisualElement => {
                if visual_target
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
                        begin_conditions.push(condition);
                        6
                    },
                    TimeConditionType::End => {
                        end_conditions.push(condition);
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
                modifiers.push(parse_time_modifier(child)?);
                8
            },
            other => {
                return Err(Error::InvalidFormat(format!(
                    "unexpected {other:?} child in SubEffectContainer"
                )));
            },
        };
        if rank < last_rank {
            return Err(Error::InvalidFormat(
                "SubEffectContainer children are not in canonical order".to_string(),
            ));
        }
        last_rank = rank;
    }
    if behavior.is_some() && kind != TimeNodeKind::Behavior {
        return Err(Error::InvalidFormat(
            "subeffect behavior requires a behavior time node".to_string(),
        ));
    }
    if visual_target.is_some() && kind != TimeNodeKind::Media {
        return Err(Error::InvalidFormat(
            "subeffect visual target requires a media time node".to_string(),
        ));
    }
    Ok(TimeSubEffect {
        atom,
        properties,
        behavior,
        visual_target,
        begin_conditions,
        end_conditions,
        modifiers,
    })
}

fn set_subeffect_behavior(
    slot: &mut Option<TimeSubEffectBehavior>,
    behavior: TimeSubEffectBehavior,
) -> Result<()> {
    if slot.replace(behavior).is_some() {
        return Err(Error::InvalidFormat(
            "SubEffectContainer contains multiple animation behaviors".to_string(),
        ));
    }
    Ok(())
}

fn set_time_node_behavior(
    slot: &mut Option<TimeNodeBehavior>,
    behavior: TimeNodeBehavior,
) -> Result<()> {
    set_once(slot, behavior, "animation behavior")
}

fn set_once<T>(slot: &mut Option<T>, value: T, field: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(Error::InvalidFormat(format!(
            "ExtTimeNode contains multiple {field} records"
        )));
    }
    Ok(())
}

/// Parse the exact 32-byte payload of a `TimeNodeAtom`.
pub fn parse_time_node_atom(record: &Record) -> Result<TimeNodeAtom> {
    require_atom(record, RecordType::TimeNode, 0, 32, "TimeNodeAtom")?;
    let flags = read_u32(&record.data, 28);
    let fill = parse_optional_time_value(
        flags & 0x01 != 0,
        read_u32(&record.data, 12),
        TimeNodeFill::parse,
        "TimeNodeAtom fill",
    )?;
    let restart = parse_optional_time_value(
        flags & 0x02 != 0,
        read_u32(&record.data, 4),
        TimeNodeRestart::parse,
        "TimeNodeAtom restart",
    )?;
    let node_type = parse_optional_time_value(
        flags & 0x08 != 0,
        read_u32(&record.data, 8),
        TimeNodeKind::parse,
        "TimeNodeAtom type",
    )?;
    let duration = i32::from_le_bytes(record.data[24..28].try_into().expect("length checked"));
    let duration_ms = if flags & 0x10 != 0 {
        Some(duration)
    } else if duration == 0 {
        None
    } else {
        return Err(Error::InvalidFormat(
            "TimeNodeAtom duration must be zero when not explicitly set".to_string(),
        ));
    };
    Ok(TimeNodeAtom {
        fill,
        restart,
        node_type,
        duration_ms,
    })
}

/// Discover PowerPoint 2002 timing and build records in `BinaryTagData`.
pub fn parse_slide_animation_extension(data: &[u8]) -> Result<SlideAnimationExtension> {
    let mut extension = SlideAnimationExtension::default();
    let mut offset = 0usize;
    let mut linked_shape_array_closed = false;
    while offset < data.len() {
        if data.len() - offset < 8 {
            return Err(Error::Corrupted(
                "slide binary tag ends with a partial record header".to_string(),
            ));
        }
        let (record, consumed) = Record::parse(data, offset)?;
        if record.data_length as usize != record.data.len() || consumed < 8 {
            return Err(Error::Corrupted(format!(
                "slide binary tag contains a truncated {:?} record",
                record.record_type
            )));
        }
        if record.record_type != RecordType::LinkedShape10Atom
            && !extension.linked_shapes.is_empty()
        {
            linked_shape_array_closed = true;
        }
        match record.record_type {
            RecordType::LinkedSlide10Atom => {
                if extension.linked_slide.is_some() {
                    return Err(Error::InvalidFormat(
                        "___PPT10 contains multiple LinkedSlide10Atom records".to_string(),
                    ));
                }
                if !extension.linked_shapes.is_empty() {
                    return Err(Error::InvalidFormat(
                        "LinkedSlide10Atom must precede its LinkedShape10Atom array".to_string(),
                    ));
                }
                extension.linked_slide = Some(LinkedSlide::parse_record(&record)?);
            },
            RecordType::LinkedShape10Atom => {
                let linked_slide = extension.linked_slide.ok_or_else(|| {
                    Error::InvalidFormat(
                        "LinkedShape10Atom requires a preceding LinkedSlide10Atom".to_string(),
                    )
                })?;
                if linked_shape_array_closed {
                    return Err(Error::InvalidFormat(
                        "LinkedShape10Atom array must be contiguous".to_string(),
                    ));
                }
                let declared_count =
                    usize::try_from(linked_slide.linked_shape_count()).map_err(|_| {
                        Error::InvalidFormat(
                            "LinkedSlide10Atom shape count does not fit this platform".to_string(),
                        )
                    })?;
                if extension.linked_shapes.len() >= declared_count {
                    return Err(Error::InvalidFormat(
                        "LinkedShape10Atom array exceeds its declared count".to_string(),
                    ));
                }
                extension
                    .linked_shapes
                    .push(LinkedShape::parse_record(&record)?);
            },
            RecordType::ExtTimeNode => {
                if extension.time_node.is_some() {
                    return Err(Error::InvalidFormat(
                        "___PPT10 contains multiple root ExtTimeNode records".to_string(),
                    ));
                }
                extension.time_node = Some(parse_extended_time_node(&record)?);
            },
            RecordType::BuildList => {
                if extension.build_list.is_some() {
                    return Err(Error::InvalidFormat(
                        "___PPT10 contains multiple BuildList records".to_string(),
                    ));
                }
                extension.build_list = Some(parse_build_list(&record)?);
            },
            RecordType::SlideFlags10Atom => {
                if extension.slide_flags.is_some() {
                    return Err(Error::InvalidFormat(
                        "___PPT10 contains multiple SlideFlags10Atom records".to_string(),
                    ));
                }
                extension.slide_flags = Some(Flags::parse_record(&record)?);
            },
            RecordType::SlideTime10Atom => {
                if extension.creation_time_filetime.is_some() {
                    return Err(Error::InvalidFormat(
                        "___PPT10 contains multiple SlideTime10Atom records".to_string(),
                    ));
                }
                extension.creation_time_filetime =
                    Some(SlideTime::parse_record(&record)?.file_time());
            },
            RecordType::HashCode10Atom => {
                if extension.animation_hash.is_some() {
                    return Err(Error::InvalidFormat(
                        "___PPT10 contains multiple HashCode10Atom records".to_string(),
                    ));
                }
                extension.animation_hash =
                    Some(crate::animation::hash::Hash10::parse_record(&record)?.hash());
            },
            _ => {},
        }
        offset = offset
            .checked_add(consumed)
            .ok_or_else(|| Error::Corrupted("slide binary tag offset overflow".to_string()))?;
    }
    if let Some(linked_slide) = extension.linked_slide {
        let declared_count = usize::try_from(linked_slide.linked_shape_count()).map_err(|_| {
            Error::InvalidFormat(
                "LinkedSlide10Atom shape count does not fit this platform".to_string(),
            )
        })?;
        if extension.linked_shapes.len() != declared_count {
            return Err(Error::InvalidFormat(format!(
                "LinkedSlide10Atom declares {declared_count} linked shapes but {} were present",
                extension.linked_shapes.len()
            )));
        }
    }
    Ok(extension)
}

/// Parse a time-node property list in its containing-node context.
pub fn parse_time_node_property_list(
    record: &Record,
    context: TimePropertyListContext,
) -> Result<TimeNodePropertyList> {
    require_container(record, RecordType::TimePropertyList, 0, "TimePropertyList")?;
    let mut seen = std::collections::HashSet::with_capacity(record.children.len());
    let mut properties = Vec::with_capacity(record.children.len());
    for child in &record.children {
        if child.record_type != RecordType::TimeVariant || child.version != 0 {
            return Err(Error::InvalidFormat(
                "invalid TimePropertyList child".to_string(),
            ));
        }
        let id = child.instance;
        if !seen.insert(id) {
            return Err(Error::InvalidFormat(format!(
                "duplicate time property {id:#X}"
            )));
        }
        let property = parse_time_node_property(child)?;
        if matches!(context, TimePropertyListContext::TimeNode) && matches!(id, 0x05 | 0x06) {
            return Err(Error::InvalidFormat(
                "subeffect-only property on time node".to_string(),
            ));
        }
        if matches!(context, TimePropertyListContext::SubEffect)
            && matches!(id, 0x09..=0x0B | 0x0F..=0x14 | 0x1A)
        {
            return Err(Error::InvalidFormat(
                "time-node-only property on subeffect".to_string(),
            ));
        }
        properties.push(property);
    }
    if properties
        .iter()
        .any(|p| matches!(p, TimeNodeProperty::EventFilter(_)))
        && !properties.iter().any(|p| {
            matches!(
                p,
                TimeNodeProperty::EffectNodeType(TimeEffectNodeType::InteractiveSequence)
            )
        })
    {
        return Err(Error::InvalidFormat(
            "event filter requires an interactive sequence".to_string(),
        ));
    }
    if !has_valid_time_effect_properties(&properties) {
        return Err(Error::InvalidFormat(
            "invalid effect ID, type, or direction combination".to_string(),
        ));
    }
    Ok(TimeNodePropertyList { properties })
}

fn parse_time_node_property(record: &Record) -> Result<TimeNodeProperty> {
    require_time_variant_payload(record)?;
    let data = &record.data;
    let int = || -> Result<i32> {
        if data.len() != 5 || data[0] != 1 {
            return Err(Error::InvalidFormat(
                "invalid integer time variant".to_string(),
            ));
        }
        Ok(i32::from_le_bytes(
            data[1..5].try_into().expect("length checked"),
        ))
    };
    let boolean = || -> Result<bool> {
        if data.len() != 2 || data[0] != 0 {
            return Err(Error::InvalidFormat(
                "invalid boolean time variant".to_string(),
            ));
        }
        parse_bool1(data[1], "TimeVariant.boolValue")
    };
    let string = || -> Result<String> {
        if data.len() < 3 || data.len() % 2 != 1 || data[0] != 3 {
            return Err(Error::InvalidFormat(
                "invalid string time variant".to_string(),
            ));
        }
        String::from_utf16(
            &data[1..]
                .chunks_exact(2)
                .map(|b| u16::from_le_bytes([b[0], b[1]]))
                .collect::<Vec<_>>(),
        )
        .map_err(|_| Error::InvalidFormat("invalid UTF-16 time variant".to_string()))
    };
    Ok(match record.instance {
        0x02 => TimeNodeProperty::DisplayHidden(match int()? {
            0 => false,
            1 => true,
            v => return Err(Error::InvalidFormat(format!("invalid display type {v}"))),
        }),
        0x05 => TimeNodeProperty::MasterRelation(match int()? {
            0 => TimeMasterRelation::DoNotStart,
            2 => TimeMasterRelation::StartWithMaster,
            v => {
                return Err(Error::InvalidFormat(format!("invalid master relation {v}")));
            },
        }),
        0x06 if int()? == 1 => TimeNodeProperty::SubType,
        0x06 => return Err(Error::InvalidFormat("invalid time subtype".to_string())),
        0x09 => TimeNodeProperty::EffectId(int()?),
        0x0A => TimeNodeProperty::EffectDirection(int()?),
        0x0B => TimeNodeProperty::EffectType(match int()? {
            1 => TimeEffectType::Entrance,
            2 => TimeEffectType::Exit,
            3 => TimeEffectType::Emphasis,
            4 => TimeEffectType::MotionPath,
            5 => TimeEffectType::ActionVerb,
            6 => TimeEffectType::MediaCommand,
            v => return Err(Error::InvalidFormat(format!("invalid effect type {v}"))),
        }),
        0x0D => TimeNodeProperty::AfterEffect(boolean()?),
        0x0F => TimeNodeProperty::SlideCount(int()?),
        0x10 => {
            let value = string()?;
            if !is_valid_time_filter(&value) {
                return Err(Error::InvalidFormat("invalid time filter".to_string()));
            }
            TimeNodeProperty::TimeFilter(value)
        },
        0x11 => {
            let value = string()?;
            if value != "cancelBubble" {
                return Err(Error::InvalidFormat("invalid event filter".to_string()));
            }
            TimeNodeProperty::EventFilter(value)
        },
        0x12 => TimeNodeProperty::HideWhenStopped(boolean()?),
        0x13 => TimeNodeProperty::GroupId(int()?),
        0x14 => TimeNodeProperty::EffectNodeType(match int()? {
            1 => TimeEffectNodeType::ClickEffect,
            2 => TimeEffectNodeType::WithPrevious,
            3 => TimeEffectNodeType::AfterPrevious,
            4 => TimeEffectNodeType::MainSequence,
            5 => TimeEffectNodeType::InteractiveSequence,
            6 => TimeEffectNodeType::ClickParallel,
            7 => TimeEffectNodeType::WithGroup,
            8 => TimeEffectNodeType::AfterGroup,
            9 => TimeEffectNodeType::TimingRoot,
            v => {
                return Err(Error::InvalidFormat(format!(
                    "invalid effect node type {v}"
                )));
            },
        }),
        0x15 => TimeNodeProperty::PlaceholderNode(boolean()?),
        0x16 => {
            if data.len() != 5 || data[0] != 2 {
                return Err(Error::InvalidFormat("invalid media volume".to_string()));
            }
            let v = f32::from_le_bytes(data[1..5].try_into().expect("length checked"));
            if !v.is_finite() || !(0.0..=100000.0).contains(&v) {
                return Err(Error::InvalidFormat(
                    "media volume out of range".to_string(),
                ));
            }
            TimeNodeProperty::MediaVolume(v)
        },
        0x17 => TimeNodeProperty::MediaMute(boolean()?),
        0x1A => TimeNodeProperty::ZoomToFullScreen(boolean()?),
        id => {
            return Err(Error::InvalidFormat(format!(
                "unknown time property {id:#X}"
            )));
        },
    })
}
