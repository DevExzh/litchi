//! Animation record parser.
//!
//! Parses PowerPoint binary animation records into structured types.

use super::triggers::IterationType;
use super::types::{
    AfterEffect, AnimationInfo, BuildAtom, BuildKind, BuildList, BuildListEntry, ChartBuild,
    ChartBuildAtom, ChartBuildType, DiagramBuild, DiagramBuildAtom, DiagramBuildType,
    ExtendedTimeNode, LegacyAnimationAtom, LegacyAnimationBuild, LegacyAnimationEffect,
    LegacyTextBuildSubEffect, ParagraphBuild, ParagraphBuildAtom, ParagraphBuildLevel,
    ParagraphBuildType, SlideAnimationExtension, TimeAnimateBehavior, TimeAnimateBehaviorAtom,
    TimeAnimateCalculationMode, TimeAnimateColor, TimeAnimateColorBy, TimeAnimateValueType,
    TimeAnimationValue, TimeAnimationValueList, TimeBehavior, TimeBehaviorAdditive,
    TimeBehaviorAtom, TimeBehaviorProperty, TimeBehaviorPropertyList, TimeColorBehavior,
    TimeColorBehaviorAtom, TimeColorDirection, TimeColorModel, TimeCommandBehavior,
    TimeCommandBehaviorAtom, TimeCommandBehaviorType, TimeCondition, TimeConditionAtom,
    TimeConditionType, TimeEffectBehavior, TimeEffectBehaviorAtom, TimeEffectFilter,
    TimeEffectNodeType, TimeEffectTransition, TimeEffectType, TimeIterateData,
    TimeIterateDirection, TimeIterateIntervalType, TimeIterateType, TimeMasterRelation,
    TimeModifier, TimeMotionBehavior, TimeMotionBehaviorAtom, TimeMotionOrigin, TimeNodeAtom,
    TimeNodeFill, TimeNodeKind, TimeNodeProperty, TimeNodePropertyList, TimeNodeRestart,
    TimePropertyListContext, TimeRotationBehavior, TimeRotationBehaviorAtom, TimeRotationDirection,
    TimeScaleBehavior, TimeScaleBehaviorAtom, TimeSequenceData, TimeSequenceNextAction,
    TimeSequencePreviousAction, TimeSetBehavior, TimeSetBehaviorAtom, TimeTriggerEvent,
    TimeTriggerObject, TimeVariantValue, TimeVisualElement, TimeVisualElementKind,
    is_valid_animation_attribute_name, is_valid_motion_path, is_valid_runtime_context,
    is_valid_time_animate_value, is_valid_time_filter, is_valid_time_formula,
    is_valid_time_points_types, is_valid_time_set_value, time_animation_attribute_value_type,
    time_set_attribute_value_type,
};
use crate::consts::PptRecordType;
use crate::ppt::package::{PptError, Result};
use crate::ppt::records::PptRecord;

/// Parse animation info from AnimationInfo container record.
pub fn parse_animation_info(record: &PptRecord) -> Result<AnimationInfo> {
    if record.record_type != PptRecordType::AnimationInfo {
        return Err(PptError::InvalidFormat(format!(
            "Expected AnimationInfo record, got {:?}",
            record.record_type
        )));
    }
    if record.version != 0x0F || record.instance != 0 {
        return Err(PptError::Corrupted(format!(
            "AnimationInfo requires version 15 and instance 0; got version {} and instance {}",
            record.version, record.instance
        )));
    }

    let mut info = AnimationInfo::new();
    let atom_record = record.children.first().ok_or_else(|| {
        PptError::Corrupted("AnimationInfo is missing its AnimationInfoAtom".to_string())
    })?;
    if atom_record.record_type != PptRecordType::AnimationInfoAtom {
        return Err(PptError::InvalidFormat(
            "AnimationInfoAtom must be the first AnimationInfo child".to_string(),
        ));
    }
    let atom = parse_animation_info_atom(atom_record)?;
    info.after_effect_color = Some(atom.dim_color);
    info.iteration = match atom.text_build_sub_effect {
        LegacyTextBuildSubEffect::AllAtOnce => IterationType::All,
        LegacyTextBuildSubEffect::ByWord => IterationType::ByWord,
        LegacyTextBuildSubEffect::ByCharacter => IterationType::ByLetter,
    };
    info.legacy_atom = Some(atom);
    for child in record.children.iter().skip(1) {
        if child.record_type == PptRecordType::AnimationInfoAtom {
            return Err(PptError::InvalidFormat(
                "AnimationInfo contains multiple AnimationInfoAtom records".to_string(),
            ));
        }
        info.raw_records.push(child.clone());
    }

    Ok(info)
}

/// Parse the exact PowerPoint 97 `AnimationInfoAtom` payload.
pub fn parse_animation_info_atom(record: &PptRecord) -> Result<LegacyAnimationAtom> {
    if record.record_type != PptRecordType::AnimationInfoAtom {
        return Err(PptError::InvalidFormat(format!(
            "Expected AnimationInfoAtom record, got {:?}",
            record.record_type
        )));
    }
    if record.version != 1 || record.instance != 0 || record.data.len() != 28 {
        return Err(PptError::Corrupted(format!(
            "AnimationInfoAtom requires version 1, instance 0, and 28 data bytes; got version {}, instance {}, length {}",
            record.version,
            record.instance,
            record.data.len()
        )));
    }

    let dim_color = u32::from_le_bytes(record.data[0..4].try_into().expect("length checked"));
    let flags = u16::from_le_bytes(record.data[4..6].try_into().expect("length checked"));
    let mut decoded_flags = [false; 8];
    for (index, decoded) in decoded_flags.iter_mut().enumerate() {
        let value = (flags >> (index * 2)) & 0x03;
        if value > 1 {
            return Err(PptError::InvalidFormat(format!(
                "AnimationInfoAtom flag field {index} has invalid bool2 value {value}"
            )));
        }
        *decoded = value == 1;
    }
    let sound_id_ref = u32::from_le_bytes(record.data[8..12].try_into().expect("length checked"));
    let delay_time_ms = i32::from_le_bytes(record.data[12..16].try_into().expect("length checked"));
    if decoded_flags[1] && delay_time_ms < 0 {
        return Err(PptError::InvalidFormat(
            "automatic AnimationInfoAtom has a negative delay".to_string(),
        ));
    }
    let order_id = i16::from_le_bytes(record.data[16..18].try_into().expect("length checked"));
    if order_id < -2 {
        return Err(PptError::InvalidFormat(format!(
            "AnimationInfoAtom orderID {order_id} is less than -2"
        )));
    }
    let slide_count = u16::from_le_bytes(record.data[18..20].try_into().expect("length checked"));
    let build_type = LegacyAnimationBuild::parse(record.data[20]).ok_or_else(|| {
        PptError::InvalidFormat(format!(
            "invalid AnimationInfoAtom animBuildType {:#04X}",
            record.data[20]
        ))
    })?;
    let effect = LegacyAnimationEffect::parse(record.data[21]).ok_or_else(|| {
        PptError::InvalidFormat(format!(
            "invalid AnimationInfoAtom animEffect {:#04X}",
            record.data[21]
        ))
    })?;
    let effect_direction = record.data[22];
    if !effect.accepts_direction(effect_direction) {
        return Err(PptError::InvalidFormat(format!(
            "AnimationInfoAtom direction {effect_direction:#04X} is invalid for {effect:?}"
        )));
    }
    let after_effect = match record.data[23] {
        0 => AfterEffect::None,
        1 => AfterEffect::DimToColor,
        2 => AfterEffect::HideOnNextClick,
        3 => AfterEffect::Hide,
        value => {
            return Err(PptError::InvalidFormat(format!(
                "invalid AnimationInfoAtom animAfterEffect {value:#04X}"
            )));
        },
    };
    let text_build_sub_effect =
        LegacyTextBuildSubEffect::parse(record.data[24]).ok_or_else(|| {
            PptError::InvalidFormat(format!(
                "invalid AnimationInfoAtom textBuildSubEffect {:#04X}",
                record.data[24]
            ))
        })?;

    Ok(LegacyAnimationAtom {
        dim_color,
        reverse: decoded_flags[0],
        automatic: decoded_flags[1],
        has_sound: decoded_flags[2],
        stop_sound: decoded_flags[3],
        play: decoded_flags[4],
        synchronous: decoded_flags[5],
        hide_while_not_playing: decoded_flags[6],
        animate_background: decoded_flags[7],
        sound_id_ref,
        delay_time_ms,
        order_id,
        slide_count,
        build_type,
        effect,
        effect_direction,
        after_effect,
        text_build_sub_effect,
        ole_verb: record.data[25],
    })
}

/// Parse the exact envelope and atom of an extended PowerPoint 2002 time node.
pub fn parse_extended_time_node(record: &PptRecord) -> Result<ExtendedTimeNode> {
    require_container(record, PptRecordType::ExtTimeNode, 1, "ExtTimeNode")?;
    let atom_record = record.children.first().ok_or_else(|| {
        PptError::Corrupted("ExtTimeNode is missing its TimeNodeAtom".to_string())
    })?;
    let atom = parse_time_node_atom(atom_record)?;
    if record.children[1..]
        .iter()
        .any(|child| child.record_type == PptRecordType::TimeNode)
    {
        return Err(PptError::InvalidFormat(
            "ExtTimeNode contains multiple TimeNodeAtom records".to_string(),
        ));
    }
    let (properties, child_start) = if record
        .children
        .get(1)
        .is_some_and(|child| child.record_type == PptRecordType::TimePropertyList)
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
    if record.children[child_start..]
        .iter()
        .any(|child| child.record_type == PptRecordType::TimePropertyList)
    {
        return Err(PptError::InvalidFormat(
            "TimePropertyList must immediately follow TimeNodeAtom".to_string(),
        ));
    }
    Ok(ExtendedTimeNode {
        atom,
        properties,
        children: record.children[child_start..].to_vec(),
    })
}

/// Parse the exact 32-byte payload of a `TimeNodeAtom`.
pub fn parse_time_node_atom(record: &PptRecord) -> Result<TimeNodeAtom> {
    require_atom(record, PptRecordType::TimeNode, 0, 32, "TimeNodeAtom")?;
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
        return Err(PptError::InvalidFormat(
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
    while offset < data.len() {
        if data.len() - offset < 8 {
            return Err(PptError::Corrupted(
                "slide binary tag ends with a partial record header".to_string(),
            ));
        }
        let (record, consumed) = PptRecord::parse(data, offset)?;
        if record.data_length as usize != record.data.len() || consumed < 8 {
            return Err(PptError::Corrupted(format!(
                "slide binary tag contains a truncated {:?} record",
                record.record_type
            )));
        }
        match record.record_type {
            PptRecordType::ExtTimeNode => {
                if extension.time_node.is_some() {
                    return Err(PptError::InvalidFormat(
                        "___PPT10 contains multiple root ExtTimeNode records".to_string(),
                    ));
                }
                extension.time_node = Some(parse_extended_time_node(&record)?);
            },
            PptRecordType::BuildList => {
                if extension.build_list.is_some() {
                    return Err(PptError::InvalidFormat(
                        "___PPT10 contains multiple BuildList records".to_string(),
                    ));
                }
                extension.build_list = Some(parse_build_list(&record)?);
            },
            _ => {},
        }
        offset = offset
            .checked_add(consumed)
            .ok_or_else(|| PptError::Corrupted("slide binary tag offset overflow".to_string()))?;
    }
    Ok(extension)
}

/// Parse a time-node property list in its containing-node context.
pub fn parse_time_node_property_list(
    record: &PptRecord,
    context: TimePropertyListContext,
) -> Result<TimeNodePropertyList> {
    require_container(
        record,
        PptRecordType::TimePropertyList,
        0,
        "TimePropertyList",
    )?;
    let mut seen = std::collections::HashSet::with_capacity(record.children.len());
    let mut properties = Vec::with_capacity(record.children.len());
    for child in &record.children {
        if child.record_type != PptRecordType::TimeVariant || child.version != 0 {
            return Err(PptError::InvalidFormat(
                "invalid TimePropertyList child".to_string(),
            ));
        }
        let id = child.instance;
        if !seen.insert(id) {
            return Err(PptError::InvalidFormat(format!(
                "duplicate time property {id:#X}"
            )));
        }
        let property = parse_time_node_property(child)?;
        if matches!(context, TimePropertyListContext::TimeNode) && matches!(id, 0x05 | 0x06) {
            return Err(PptError::InvalidFormat(
                "subeffect-only property on time node".to_string(),
            ));
        }
        if matches!(context, TimePropertyListContext::SubEffect)
            && matches!(id, 0x09..=0x0B | 0x0F..=0x14 | 0x1A)
        {
            return Err(PptError::InvalidFormat(
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
        return Err(PptError::InvalidFormat(
            "event filter requires an interactive sequence".to_string(),
        ));
    }
    Ok(TimeNodePropertyList { properties })
}

fn parse_time_node_property(record: &PptRecord) -> Result<TimeNodeProperty> {
    require_time_variant_payload(record)?;
    let data = &record.data;
    let int = || -> Result<i32> {
        if data.len() != 5 || data[0] != 1 {
            return Err(PptError::InvalidFormat(
                "invalid integer time variant".to_string(),
            ));
        }
        Ok(i32::from_le_bytes(
            data[1..5].try_into().expect("length checked"),
        ))
    };
    let boolean = || -> Result<bool> {
        if data.len() != 2 || data[0] != 0 {
            return Err(PptError::InvalidFormat(
                "invalid boolean time variant".to_string(),
            ));
        }
        parse_bool1(data[1], "TimeVariant.boolValue")
    };
    let string = || -> Result<String> {
        if data.len() < 3 || data.len() % 2 != 1 || data[0] != 3 {
            return Err(PptError::InvalidFormat(
                "invalid string time variant".to_string(),
            ));
        }
        String::from_utf16(
            &data[1..]
                .chunks_exact(2)
                .map(|b| u16::from_le_bytes([b[0], b[1]]))
                .collect::<Vec<_>>(),
        )
        .map_err(|_| PptError::InvalidFormat("invalid UTF-16 time variant".to_string()))
    };
    Ok(match record.instance {
        0x02 => TimeNodeProperty::DisplayHidden(match int()? {
            0 => false,
            1 => true,
            v => return Err(PptError::InvalidFormat(format!("invalid display type {v}"))),
        }),
        0x05 => TimeNodeProperty::MasterRelation(match int()? {
            0 => TimeMasterRelation::DoNotStart,
            2 => TimeMasterRelation::StartWithMaster,
            v => {
                return Err(PptError::InvalidFormat(format!(
                    "invalid master relation {v}"
                )));
            },
        }),
        0x06 if int()? == 1 => TimeNodeProperty::SubType,
        0x06 => return Err(PptError::InvalidFormat("invalid time subtype".to_string())),
        0x09 => TimeNodeProperty::EffectId(int()?),
        0x0A => TimeNodeProperty::EffectDirection(int()?),
        0x0B => TimeNodeProperty::EffectType(match int()? {
            1 => TimeEffectType::Entrance,
            2 => TimeEffectType::Exit,
            3 => TimeEffectType::Emphasis,
            4 => TimeEffectType::MotionPath,
            5 => TimeEffectType::ActionVerb,
            6 => TimeEffectType::MediaCommand,
            v => return Err(PptError::InvalidFormat(format!("invalid effect type {v}"))),
        }),
        0x0D => TimeNodeProperty::AfterEffect(boolean()?),
        0x0F => TimeNodeProperty::SlideCount(int()?),
        0x10 => {
            let value = string()?;
            if !is_valid_time_filter(&value) {
                return Err(PptError::InvalidFormat("invalid time filter".to_string()));
            }
            TimeNodeProperty::TimeFilter(value)
        },
        0x11 => {
            let value = string()?;
            if value != "cancelBubble" {
                return Err(PptError::InvalidFormat("invalid event filter".to_string()));
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
                return Err(PptError::InvalidFormat(format!(
                    "invalid effect node type {v}"
                )));
            },
        }),
        0x15 => TimeNodeProperty::PlaceholderNode(boolean()?),
        0x16 => {
            if data.len() != 5 || data[0] != 2 {
                return Err(PptError::InvalidFormat("invalid media volume".to_string()));
            }
            let v = f32::from_le_bytes(data[1..5].try_into().expect("length checked"));
            if !v.is_finite() || !(0.0..=100000.0).contains(&v) {
                return Err(PptError::InvalidFormat(
                    "media volume out of range".to_string(),
                ));
            }
            TimeNodeProperty::MediaVolume(v)
        },
        0x17 => TimeNodeProperty::MediaMute(boolean()?),
        0x1A => TimeNodeProperty::ZoomToFullScreen(boolean()?),
        id => {
            return Err(PptError::InvalidFormat(format!(
                "unknown time property {id:#X}"
            )));
        },
    })
}

/// Parse the common behavior information shared by all extended animation behaviors.
pub fn parse_time_behavior(record: &PptRecord) -> Result<TimeBehavior> {
    require_container(
        record,
        PptRecordType::TimeBehaviorContainer,
        0,
        "TimeBehaviorContainer",
    )?;
    let atom_record = record
        .children
        .first()
        .ok_or_else(|| PptError::InvalidFormat("TimeBehaviorContainer has no atom".to_string()))?;
    let atom = parse_time_behavior_atom(atom_record)?;
    let mut index = 1;
    let attribute_names = if record
        .children
        .get(index)
        .is_some_and(|child| child.record_type == PptRecordType::TimeVariantList)
    {
        let names = parse_time_string_list(&record.children[index])?;
        index += 1;
        Some(names)
    } else {
        None
    };
    let properties = if record
        .children
        .get(index)
        .is_some_and(|child| child.record_type == PptRecordType::TimePropertyList)
    {
        let properties = parse_time_behavior_property_list(&record.children[index])?;
        index += 1;
        Some(properties)
    } else {
        None
    };
    let target = record
        .children
        .get(index)
        .ok_or_else(|| PptError::InvalidFormat("TimeBehaviorContainer has no target".to_string()))
        .and_then(parse_time_visual_element)?;
    index += 1;
    if index != record.children.len() {
        return Err(PptError::InvalidFormat(
            "TimeBehaviorContainer has invalid child order or extra children".to_string(),
        ));
    }
    Ok(TimeBehavior {
        atom,
        attribute_names,
        properties,
        target,
    })
}

/// Parse an exact 16-byte `TimeBehaviorAtom` payload.
pub fn parse_time_behavior_atom(record: &PptRecord) -> Result<TimeBehaviorAtom> {
    require_atom(
        record,
        PptRecordType::TimeBehavior,
        0,
        16,
        "TimeBehaviorAtom",
    )?;
    let flags = read_u32(&record.data, 0);
    let additive_value = read_u32(&record.data, 4);
    let additive = if flags & 0x01 != 0 {
        Some(match additive_value {
            0 => TimeBehaviorAdditive::Override,
            1 => TimeBehaviorAdditive::Add,
            value => {
                return Err(PptError::InvalidFormat(format!(
                    "invalid TimeBehavior additive mode {value}"
                )));
            },
        })
    } else if additive_value == 0 {
        None
    } else {
        return Err(PptError::InvalidFormat(
            "TimeBehavior additive mode must be zero when not explicitly set".to_string(),
        ));
    };
    if read_u32(&record.data, 8) != 0 || read_u32(&record.data, 12) != 0 {
        return Err(PptError::InvalidFormat(
            "TimeBehavior accumulation and transform modes must be zero".to_string(),
        ));
    }
    Ok(TimeBehaviorAtom {
        additive,
        attribute_names_used: flags & 0x04 != 0,
    })
}

/// Parse an exact generic property-animation behavior container.
pub fn parse_time_animate_behavior(record: &PptRecord) -> Result<TimeAnimateBehavior> {
    require_container(
        record,
        PptRecordType::TimeAnimateBehaviorContainer,
        0,
        "TimeAnimateBehaviorContainer",
    )?;
    let atom = record
        .children
        .first()
        .ok_or_else(|| PptError::InvalidFormat("animate behavior has no atom".to_string()))
        .and_then(parse_time_animate_behavior_atom)?;
    let mut index = 1;
    let values = if record
        .children
        .get(index)
        .is_some_and(|child| child.record_type == PptRecordType::TimeAnimationValueList)
    {
        let values = parse_time_animation_value_list(&record.children[index])?;
        index += 1;
        Some(values)
    } else {
        None
    };
    let mut take_string = |instance| -> Result<Option<String>> {
        if record.children.get(index).is_some_and(|child| {
            child.record_type == PptRecordType::TimeVariant && child.instance == instance
        }) {
            let value = parse_time_variant_string(&record.children[index])?;
            index += 1;
            Ok(Some(value))
        } else {
            Ok(None)
        }
    };
    let by = take_string(1)?;
    let from = take_string(2)?;
    let to = take_string(3)?;
    let behavior = record
        .children
        .get(index)
        .ok_or_else(|| PptError::InvalidFormat("animate behavior has no target".to_string()))
        .and_then(parse_time_behavior)?;
    index += 1;
    if index != record.children.len() {
        return Err(PptError::InvalidFormat(
            "animate behavior has invalid child order or extra children".to_string(),
        ));
    }
    let animate = TimeAnimateBehavior {
        atom,
        values,
        by,
        from,
        to,
        behavior,
    };
    validate_animate_behavior(&animate)?;
    Ok(animate)
}

/// Parse an exact 12-byte `TimeAnimateBehaviorAtom` payload.
pub fn parse_time_animate_behavior_atom(record: &PptRecord) -> Result<TimeAnimateBehaviorAtom> {
    require_atom(
        record,
        PptRecordType::TimeAnimateBehavior,
        0,
        12,
        "TimeAnimateBehaviorAtom",
    )?;
    let mode_value = read_u32(&record.data, 0);
    let flags = read_u32(&record.data, 4);
    let calculation_mode = if flags & 0x08 != 0 {
        Some(match mode_value {
            0 => TimeAnimateCalculationMode::Discrete,
            1 => TimeAnimateCalculationMode::Linear,
            2 => TimeAnimateCalculationMode::Formula,
            value => {
                return Err(PptError::InvalidFormat(format!(
                    "invalid animate calculation mode {value}"
                )));
            },
        })
    } else if mode_value == 1 {
        None
    } else {
        return Err(PptError::InvalidFormat(
            "animate calculation mode must be linear when unused".to_string(),
        ));
    };
    let type_value = read_u32(&record.data, 8);
    let value_type = if flags & 0x20 != 0 {
        Some(match type_value {
            0 => TimeAnimateValueType::String,
            1 => TimeAnimateValueType::Number,
            2 => TimeAnimateValueType::Color,
            value => {
                return Err(PptError::InvalidFormat(format!(
                    "invalid animate value type {value}"
                )));
            },
        })
    } else if type_value == 1 {
        None
    } else {
        return Err(PptError::InvalidFormat(
            "animate value type must be numeric when unused".to_string(),
        ));
    };
    Ok(TimeAnimateBehaviorAtom {
        calculation_mode,
        by_used: flags & 0x01 != 0,
        from_used: flags & 0x02 != 0,
        to_used: flags & 0x04 != 0,
        animation_values_used: flags & 0x10 != 0,
        value_type,
    })
}

/// Parse a generic animation keyframe list.
pub fn parse_time_animation_value_list(record: &PptRecord) -> Result<TimeAnimationValueList> {
    require_container(
        record,
        PptRecordType::TimeAnimationValueList,
        0,
        "TimeAnimationValueListContainer",
    )?;
    let mut entries = Vec::new();
    let mut index = 0;
    while index < record.children.len() {
        let time = parse_time_animation_value_atom(&record.children[index])?;
        index += 1;
        let value = if record.children.get(index).is_some_and(|child| {
            child.record_type == PptRecordType::TimeVariant && child.instance == 0
        }) {
            let value = parse_generic_time_variant(&record.children[index])?;
            index += 1;
            Some(value)
        } else {
            None
        };
        let formula = if record.children.get(index).is_some_and(|child| {
            child.record_type == PptRecordType::TimeVariant && child.instance == 1
        }) {
            let formula = parse_time_variant_string(&record.children[index])?;
            if !is_valid_time_formula(&formula) {
                return Err(PptError::InvalidFormat(
                    "invalid animation keyframe formula".to_string(),
                ));
            }
            index += 1;
            Some(formula)
        } else {
            None
        };
        entries.push(TimeAnimationValue {
            time,
            value,
            formula,
        });
    }
    Ok(TimeAnimationValueList { entries })
}

/// Parse an exact 4-byte `TimeAnimationValueAtom` payload.
pub fn parse_time_animation_value_atom(record: &PptRecord) -> Result<i32> {
    require_atom(
        record,
        PptRecordType::TimeAnimationValue,
        0,
        4,
        "TimeAnimationValueAtom",
    )?;
    let time = read_i32(&record.data, 0);
    if time != -1000 && !(0..=1000).contains(&time) {
        return Err(PptError::InvalidFormat(
            "animation keyframe time is out of range".to_string(),
        ));
    }
    Ok(time)
}

fn parse_generic_time_variant(record: &PptRecord) -> Result<TimeVariantValue> {
    match record.data.first() {
        Some(0) => parse_time_variant_bool(record).map(TimeVariantValue::Boolean),
        Some(1) => parse_time_variant_i32(record).map(TimeVariantValue::Integer),
        Some(2) => parse_time_variant_f32(record).map(TimeVariantValue::Float),
        Some(3) => parse_time_variant_string(record).map(TimeVariantValue::String),
        _ => Err(PptError::InvalidFormat(
            "invalid animation keyframe value type".to_string(),
        )),
    }
}

fn validate_animate_behavior(animate: &TimeAnimateBehavior) -> Result<()> {
    for (used, value, field) in [
        (animate.atom.by_used, animate.by.as_ref(), "by"),
        (animate.atom.from_used, animate.from.as_ref(), "from"),
        (animate.atom.to_used, animate.to.as_ref(), "to"),
    ] {
        if used && value.is_none() {
            return Err(PptError::InvalidFormat(format!(
                "animate {field}-use flag requires a value"
            )));
        }
    }
    if animate.atom.animation_values_used && animate.values.is_none() {
        return Err(PptError::InvalidFormat(
            "animate-values-use flag requires a keyframe list".to_string(),
        ));
    }
    if animate.from.is_some() && animate.by.is_none() && animate.to.is_none() {
        return Err(PptError::InvalidFormat(
            "animate from value requires a by or to value".to_string(),
        ));
    }
    if !animate.behavior.atom.attribute_names_used {
        return Err(PptError::InvalidFormat(
            "animate behavior requires an explicit attribute name".to_string(),
        ));
    }
    let attribute = match animate.behavior.attribute_names.as_deref() {
        Some([attribute]) => attribute.as_str(),
        _ => {
            return Err(PptError::InvalidFormat(
                "animate behavior requires exactly one attribute name".to_string(),
            ));
        },
    };
    let expected_type = time_animation_attribute_value_type(attribute).ok_or_else(|| {
        PptError::InvalidFormat(format!(
            "unsupported animate behavior attribute {attribute}"
        ))
    })?;
    let actual_type = animate
        .atom
        .value_type
        .unwrap_or(TimeAnimateValueType::Number);
    if actual_type != expected_type {
        return Err(PptError::InvalidFormat(
            "animate value type does not match its attribute".to_string(),
        ));
    }
    if [&animate.by, &animate.from, &animate.to]
        .into_iter()
        .flatten()
        .any(|value| !is_valid_time_animate_value(attribute, actual_type, value))
    {
        return Err(PptError::InvalidFormat(
            "animate value is invalid for its attribute".to_string(),
        ));
    }
    if animate.atom.calculation_mode == Some(TimeAnimateCalculationMode::Formula)
        && !animate
            .values
            .as_ref()
            .is_some_and(|list| list.entries.iter().any(|entry| entry.formula.is_some()))
    {
        return Err(PptError::InvalidFormat(
            "formula calculation mode requires a keyframe formula".to_string(),
        ));
    }
    validate_basic_behavior_properties(&animate.behavior)
}

/// Parse an exact color behavior container.
pub fn parse_time_color_behavior(record: &PptRecord) -> Result<TimeColorBehavior> {
    require_container(
        record,
        PptRecordType::TimeColorBehaviorContainer,
        0,
        "TimeColorBehaviorContainer",
    )?;
    if record.children.len() != 2 {
        return Err(PptError::InvalidFormat(
            "TimeColorBehaviorContainer requires an atom and common behavior".to_string(),
        ));
    }
    let atom = parse_time_color_behavior_atom(&record.children[0])?;
    let behavior = parse_time_behavior(&record.children[1])?;
    validate_color_behavior(&atom, &behavior)?;
    Ok(TimeColorBehavior { atom, behavior })
}

/// Parse an exact 52-byte `TimeColorBehaviorAtom` payload.
pub fn parse_time_color_behavior_atom(record: &PptRecord) -> Result<TimeColorBehaviorAtom> {
    require_atom(
        record,
        PptRecordType::TimeColorBehavior,
        0,
        52,
        "TimeColorBehaviorAtom",
    )?;
    let flags = read_u32(&record.data, 0);
    let by = (flags & 0x01 != 0)
        .then(|| parse_animate_color_by(&record.data[4..20]))
        .transpose()?;
    let from = (flags & 0x02 != 0)
        .then(|| parse_animate_color(&record.data[20..36]))
        .transpose()?;
    let to = (flags & 0x04 != 0)
        .then(|| parse_animate_color(&record.data[36..52]))
        .transpose()?;
    if from.is_some() && by.is_none() && to.is_none() {
        return Err(PptError::InvalidFormat(
            "color from value requires a by or to value".to_string(),
        ));
    }
    Ok(TimeColorBehaviorAtom {
        by,
        from,
        to,
        color_space_used: flags & 0x08 != 0,
        direction_used: flags & 0x10 != 0,
    })
}

fn parse_animate_color_by(data: &[u8]) -> Result<TimeAnimateColorBy> {
    match read_u32(data, 0) {
        0 | 1 => {
            let values = (read_i32(data, 4), read_i32(data, 8), read_i32(data, 12));
            if [values.0, values.1, values.2]
                .iter()
                .any(|value| !(-255..=255).contains(value))
            {
                return Err(PptError::InvalidFormat(
                    "color offset component is out of range".to_string(),
                ));
            }
            if read_u32(data, 0) == 0 {
                Ok(TimeAnimateColorBy::Rgb {
                    red: values.0,
                    green: values.1,
                    blue: values.2,
                })
            } else {
                Ok(TimeAnimateColorBy::Hsl {
                    hue: values.0,
                    saturation: values.1,
                    luminance: values.2,
                })
            }
        },
        2 => parse_scheme_color(data).map(TimeAnimateColorBy::Scheme),
        model => Err(PptError::InvalidFormat(format!(
            "invalid color-by model {model}"
        ))),
    }
}

fn parse_animate_color(data: &[u8]) -> Result<TimeAnimateColor> {
    match read_u32(data, 0) {
        0 => {
            let (red, green, blue) = (read_u32(data, 4), read_u32(data, 8), read_u32(data, 12));
            if red > 255 || green > 255 || blue > 255 {
                return Err(PptError::InvalidFormat(
                    "RGB color component is out of range".to_string(),
                ));
            }
            Ok(TimeAnimateColor::Rgb { red, green, blue })
        },
        2 => parse_scheme_color(data).map(TimeAnimateColor::Scheme),
        model => Err(PptError::InvalidFormat(format!(
            "invalid absolute color model {model}"
        ))),
    }
}

fn parse_scheme_color(data: &[u8]) -> Result<u32> {
    let index = read_u32(data, 4);
    if index > 7 {
        return Err(PptError::InvalidFormat(
            "scheme color index is out of range".to_string(),
        ));
    }
    Ok(index)
}

fn validate_color_behavior(atom: &TimeColorBehaviorAtom, behavior: &TimeBehavior) -> Result<()> {
    const NAMES: &[&str] = &[
        "ppt_c",
        "style.color",
        "imageData.chromakey",
        "fill.color",
        "fill.color2",
        "stroke.color",
        "stroke.color2",
        "shadow.color",
        "shadow.color2",
        "extrusion.color",
        "fillcolor",
    ];
    if !behavior.atom.attribute_names_used
        || !matches!(behavior.attribute_names.as_deref(), Some([name]) if NAMES.contains(&name.as_str()))
    {
        return Err(PptError::InvalidFormat(
            "color behavior requires exactly one supported color attribute".to_string(),
        ));
    }
    let properties = behavior
        .properties
        .as_ref()
        .map_or(&[][..], |list| list.properties.as_slice());
    if properties.iter().any(|property| {
        matches!(
            property,
            TimeBehaviorProperty::MotionPathEditRelative(_)
                | TimeBehaviorProperty::PathEditRotationAngle(_)
                | TimeBehaviorProperty::PathEditRotationX(_)
                | TimeBehaviorProperty::PathEditRotationY(_)
                | TimeBehaviorProperty::PointsTypes(_)
        )
    }) {
        return Err(PptError::InvalidFormat(
            "color behavior contains a motion-only property".to_string(),
        ));
    }
    if atom.color_space_used
        && !properties
            .iter()
            .any(|property| matches!(property, TimeBehaviorProperty::ColorModel(_)))
    {
        return Err(PptError::InvalidFormat(
            "color-space-used flag requires a color model property".to_string(),
        ));
    }
    if atom.direction_used
        && !properties
            .iter()
            .any(|property| matches!(property, TimeBehaviorProperty::ColorDirection(_)))
    {
        return Err(PptError::InvalidFormat(
            "direction-used flag requires a color direction property".to_string(),
        ));
    }
    Ok(())
}

/// Parse an exact image-effect behavior container.
pub fn parse_time_effect_behavior(record: &PptRecord) -> Result<TimeEffectBehavior> {
    require_container(
        record,
        PptRecordType::TimeEffectBehaviorContainer,
        0,
        "TimeEffectBehaviorContainer",
    )?;
    let atom = record
        .children
        .first()
        .ok_or_else(|| PptError::InvalidFormat("effect behavior has no atom".to_string()))
        .and_then(parse_time_effect_behavior_atom)?;
    let mut index = 1;
    let filter =
        if record.children.get(index).is_some_and(|child| {
            child.record_type == PptRecordType::TimeVariant && child.instance == 1
        }) {
            let value = parse_time_variant_string(&record.children[index])?;
            index += 1;
            Some(TimeEffectFilter::parse(&value).ok_or_else(|| {
                PptError::InvalidFormat(format!("invalid image-effect filter {value}"))
            })?)
        } else {
            None
        };
    let progress =
        if record.children.get(index).is_some_and(|child| {
            child.record_type == PptRecordType::TimeVariant && child.instance == 2
        }) {
            let value = parse_time_variant_f32(&record.children[index])?;
            index += 1;
            Some(value)
        } else {
            None
        };
    let runtime_context =
        if record.children.get(index).is_some_and(|child| {
            child.record_type == PptRecordType::TimeVariant && child.instance == 3
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
        .ok_or_else(|| PptError::InvalidFormat("effect behavior has no target".to_string()))
        .and_then(parse_time_behavior)?;
    index += 1;
    if index != record.children.len() {
        return Err(PptError::InvalidFormat(
            "effect behavior has invalid child order or extra children".to_string(),
        ));
    }
    let effect = TimeEffectBehavior {
        atom,
        filter,
        progress,
        runtime_context,
        behavior,
    };
    validate_effect_behavior(&effect)?;
    Ok(effect)
}

/// Parse an exact 8-byte `TimeEffectBehaviorAtom` payload.
pub fn parse_time_effect_behavior_atom(record: &PptRecord) -> Result<TimeEffectBehaviorAtom> {
    require_atom(
        record,
        PptRecordType::TimeEffectBehavior,
        0,
        8,
        "TimeEffectBehaviorAtom",
    )?;
    let flags = read_u32(&record.data, 0);
    let value = read_u32(&record.data, 4);
    let transition = if flags & 0x01 != 0 {
        Some(match value {
            0 => TimeEffectTransition::In,
            1 => TimeEffectTransition::Out,
            value => {
                return Err(PptError::InvalidFormat(format!(
                    "invalid image-effect transition {value}"
                )));
            },
        })
    } else if value == 0 {
        None
    } else {
        return Err(PptError::InvalidFormat(
            "image-effect transition must be transition-in when unused".to_string(),
        ));
    };
    Ok(TimeEffectBehaviorAtom {
        transition,
        filter_used: flags & 0x02 != 0,
        progress_used: flags & 0x04 != 0,
        runtime_context_used: flags & 0x08 != 0,
    })
}

fn validate_effect_behavior(effect: &TimeEffectBehavior) -> Result<()> {
    if effect.atom.filter_used && effect.filter.is_none() {
        return Err(PptError::InvalidFormat(
            "image-effect filter-use flag requires a filter".to_string(),
        ));
    }
    if effect.atom.progress_used && effect.progress.is_none() {
        return Err(PptError::InvalidFormat(
            "image-effect progress-use flag requires progress".to_string(),
        ));
    }
    if effect.atom.runtime_context_used && effect.runtime_context.is_none() {
        return Err(PptError::InvalidFormat(
            "image-effect runtime-context-use flag requires a runtime context".to_string(),
        ));
    }
    if effect
        .progress
        .is_some_and(|value| !(0.0..=1.0).contains(&value))
    {
        return Err(PptError::InvalidFormat(
            "image-effect progress must be between zero and one".to_string(),
        ));
    }
    if effect
        .runtime_context
        .as_deref()
        .is_some_and(|value| !is_valid_runtime_context(value))
    {
        return Err(PptError::InvalidFormat(
            "invalid image-effect runtime context".to_string(),
        ));
    }
    validate_basic_behavior_properties(&effect.behavior)
}

/// Parse an exact motion-path behavior container.
pub fn parse_time_motion_behavior(record: &PptRecord) -> Result<TimeMotionBehavior> {
    require_container(
        record,
        PptRecordType::TimeMotionBehaviorContainer,
        0,
        "TimeMotionBehaviorContainer",
    )?;
    let atom = record
        .children
        .first()
        .ok_or_else(|| PptError::InvalidFormat("motion behavior has no atom".to_string()))
        .and_then(parse_time_motion_behavior_atom)?;
    let mut index = 1;
    let path =
        if record.children.get(index).is_some_and(|child| {
            child.record_type == PptRecordType::TimeVariant && child.instance == 1
        }) {
            let value = parse_time_variant_string(&record.children[index])?;
            index += 1;
            Some(value)
        } else {
            None
        };
    let reserved =
        if record.children.get(index).is_some_and(|child| {
            child.record_type == PptRecordType::TimeVariant && child.instance == 2
        }) {
            let value = parse_time_variant_i32(&record.children[index])?;
            index += 1;
            Some(value)
        } else {
            None
        };
    let behavior = record
        .children
        .get(index)
        .ok_or_else(|| PptError::InvalidFormat("motion behavior has no target".to_string()))
        .and_then(parse_time_behavior)?;
    index += 1;
    if index != record.children.len() {
        return Err(PptError::InvalidFormat(
            "motion behavior has invalid child order or extra children".to_string(),
        ));
    }
    let motion = TimeMotionBehavior {
        atom,
        path,
        reserved,
        behavior,
    };
    validate_motion_behavior(&motion)?;
    Ok(motion)
}

/// Parse an exact 32-byte `TimeMotionBehaviorAtom` payload.
pub fn parse_time_motion_behavior_atom(record: &PptRecord) -> Result<TimeMotionBehaviorAtom> {
    require_atom(
        record,
        PptRecordType::TimeMotionBehavior,
        0,
        32,
        "TimeMotionBehaviorAtom",
    )?;
    let flags = read_u32(&record.data, 0);
    let by = (flags & 0x01 != 0).then(|| (read_f32(&record.data, 4), read_f32(&record.data, 8)));
    let from =
        (flags & 0x02 != 0).then(|| (read_f32(&record.data, 12), read_f32(&record.data, 16)));
    let to = (flags & 0x04 != 0).then(|| (read_f32(&record.data, 20), read_f32(&record.data, 24)));
    if from.is_some() && by.is_none() && to.is_none() {
        return Err(PptError::InvalidFormat(
            "motion from values require by or to values".to_string(),
        ));
    }
    let origin_value = read_u32(&record.data, 28);
    let origin = if flags & 0x08 != 0 {
        Some(match origin_value {
            0 => TimeMotionOrigin::Slide,
            1 => TimeMotionOrigin::SlideLegacy,
            2 => TimeMotionOrigin::ObjectCenter,
            value => {
                return Err(PptError::InvalidFormat(format!(
                    "invalid motion origin {value}"
                )));
            },
        })
    } else if origin_value == 2 {
        None
    } else {
        return Err(PptError::InvalidFormat(
            "motion origin must be object center when unused".to_string(),
        ));
    };
    Ok(TimeMotionBehaviorAtom {
        by,
        from,
        to,
        origin,
        path_used: flags & 0x10 != 0,
        edit_rotation_used: flags & 0x40 != 0,
        points_types_used: flags & 0x80 != 0,
    })
}

fn validate_motion_behavior(motion: &TimeMotionBehavior) -> Result<()> {
    if motion.atom.path_used && motion.path.is_none() {
        return Err(PptError::InvalidFormat(
            "motion path-use flag requires a path".to_string(),
        ));
    }
    if motion
        .path
        .as_deref()
        .is_some_and(|path| !is_valid_motion_path(path))
    {
        return Err(PptError::InvalidFormat(
            "invalid motion path syntax".to_string(),
        ));
    }
    let properties = motion
        .behavior
        .properties
        .as_ref()
        .map_or(&[][..], |list| list.properties.as_slice());
    if properties.iter().any(|property| {
        matches!(
            property,
            TimeBehaviorProperty::ColorModel(_) | TimeBehaviorProperty::ColorDirection(_)
        )
    }) {
        return Err(PptError::InvalidFormat(
            "motion behavior contains a color-only property".to_string(),
        ));
    }
    let has_angle = properties
        .iter()
        .any(|property| matches!(property, TimeBehaviorProperty::PathEditRotationAngle(_)));
    let has_x = properties
        .iter()
        .any(|property| matches!(property, TimeBehaviorProperty::PathEditRotationX(_)));
    let has_y = properties
        .iter()
        .any(|property| matches!(property, TimeBehaviorProperty::PathEditRotationY(_)));
    if motion.atom.edit_rotation_used && !(has_angle && has_x && has_y) {
        return Err(PptError::InvalidFormat(
            "motion edit-rotation flag requires angle, X, and Y properties".to_string(),
        ));
    }
    if motion.atom.points_types_used
        && !properties
            .iter()
            .any(|property| matches!(property, TimeBehaviorProperty::PointsTypes(_)))
    {
        return Err(PptError::InvalidFormat(
            "motion points-types flag requires a points-types property".to_string(),
        ));
    }
    if motion.behavior.atom.attribute_names_used
        && motion
            .behavior
            .attribute_names
            .as_ref()
            .is_some_and(|names| {
                names.len() > 2
                    || names
                        .iter()
                        .any(|name| !is_valid_animation_attribute_name(name))
            })
    {
        return Err(PptError::InvalidFormat(
            "motion behavior requires at most two supported attribute names".to_string(),
        ));
    }
    Ok(())
}

/// Parse a `TimePropertyList4TimeBehavior` record.
pub fn parse_time_behavior_property_list(record: &PptRecord) -> Result<TimeBehaviorPropertyList> {
    require_container(
        record,
        PptRecordType::TimePropertyList,
        0,
        "TimePropertyList4TimeBehavior",
    )?;
    let mut seen = std::collections::HashSet::with_capacity(record.children.len());
    let mut properties = Vec::with_capacity(record.children.len());
    for child in &record.children {
        if child.record_type != PptRecordType::TimeVariant || child.version != 0 {
            return Err(PptError::InvalidFormat(
                "invalid TimePropertyList4TimeBehavior child".to_string(),
            ));
        }
        if !seen.insert(child.instance) {
            return Err(PptError::InvalidFormat(format!(
                "duplicate time behavior property {:#X}",
                child.instance
            )));
        }
        properties.push(parse_time_behavior_property(child)?);
    }
    Ok(TimeBehaviorPropertyList { properties })
}

fn parse_time_behavior_property(record: &PptRecord) -> Result<TimeBehaviorProperty> {
    let property = match record.instance {
        0x01 => TimeBehaviorProperty::UnknownPropertyList(parse_time_variant_string(record)?),
        0x02 => {
            let value = parse_time_variant_string(record)?;
            if !is_valid_runtime_context(&value) {
                return Err(PptError::InvalidFormat(
                    "invalid time runtime context".to_string(),
                ));
            }
            TimeBehaviorProperty::RuntimeContext(value)
        },
        0x03 => TimeBehaviorProperty::MotionPathEditRelative(parse_time_variant_bool(record)?),
        0x04 => TimeBehaviorProperty::ColorModel(match parse_time_variant_i32(record)? {
            0 => TimeColorModel::Rgb,
            1 => TimeColorModel::Hsl,
            2 => TimeColorModel::Scheme,
            value => {
                return Err(PptError::InvalidFormat(format!(
                    "invalid time color model {value}"
                )));
            },
        }),
        0x05 => TimeBehaviorProperty::ColorDirection(match parse_time_variant_i32(record)? {
            0 => TimeColorDirection::Clockwise,
            1 => TimeColorDirection::CounterClockwise,
            value => {
                return Err(PptError::InvalidFormat(format!(
                    "invalid time color direction {value}"
                )));
            },
        }),
        0x06 => match parse_time_variant_i32(record)? {
            1 => TimeBehaviorProperty::Override,
            _ => {
                return Err(PptError::InvalidFormat(
                    "invalid time behavior override".to_string(),
                ));
            },
        },
        0x07 => TimeBehaviorProperty::PathEditRotationAngle(parse_time_variant_f32(record)?),
        0x08 => TimeBehaviorProperty::PathEditRotationX(parse_time_variant_f32(record)?),
        0x09 => TimeBehaviorProperty::PathEditRotationY(parse_time_variant_f32(record)?),
        0x0A => {
            let value = parse_time_variant_string(record)?;
            if !is_valid_time_points_types(&value) {
                return Err(PptError::InvalidFormat(
                    "invalid time path point types".to_string(),
                ));
            }
            TimeBehaviorProperty::PointsTypes(value)
        },
        id => {
            return Err(PptError::InvalidFormat(format!(
                "unknown time behavior property {id:#X}"
            )));
        },
    };
    Ok(property)
}

fn parse_time_string_list(record: &PptRecord) -> Result<Vec<String>> {
    require_container(
        record,
        PptRecordType::TimeVariantList,
        1,
        "TimeStringListContainer",
    )?;
    record
        .children
        .iter()
        .map(|child| {
            if child.record_type != PptRecordType::TimeVariant || child.version != 0 {
                return Err(PptError::InvalidFormat(
                    "invalid TimeStringListContainer child".to_string(),
                ));
            }
            parse_time_variant_string(child)
        })
        .collect()
}

/// Parse a `ClientVisualElementContainer` animation target.
pub fn parse_time_visual_element(record: &PptRecord) -> Result<TimeVisualElement> {
    require_container(
        record,
        PptRecordType::TimeClientVisualElement,
        0,
        "ClientVisualElementContainer",
    )?;
    if record.children.len() != 1 {
        return Err(PptError::InvalidFormat(
            "ClientVisualElementContainer requires exactly one atom".to_string(),
        ));
    }
    let atom = &record.children[0];
    if atom.record_type == PptRecordType::VisualPageAtom {
        require_atom(atom, PptRecordType::VisualPageAtom, 0, 4, "VisualPageAtom")?;
        if read_u32(&atom.data, 0) != TimeVisualElementKind::Page.as_u32() {
            return Err(PptError::InvalidFormat(
                "VisualPageAtom has a non-page target type".to_string(),
            ));
        }
        return Ok(TimeVisualElement::Page);
    }
    require_atom(
        atom,
        PptRecordType::VisualShapeAtom,
        0,
        20,
        "VisualShapeOrSoundAtom",
    )?;
    let kind = TimeVisualElementKind::parse(read_u32(&atom.data, 0))
        .ok_or_else(|| PptError::InvalidFormat("invalid visual element target type".to_string()))?;
    if kind == TimeVisualElementKind::Page {
        return Err(PptError::InvalidFormat(
            "VisualShapeOrSoundAtom cannot target a page".to_string(),
        ));
    }
    match read_u32(&atom.data, 4) {
        1 if kind == TimeVisualElementKind::ChartElement => {
            let build_type = ChartBuildType::parse(read_u32(&atom.data, 12)).ok_or_else(|| {
                PptError::InvalidFormat("invalid chart target build type".to_string())
            })?;
            let element_index = read_i32(&atom.data, 16);
            if element_index < -1 {
                return Err(PptError::InvalidFormat(
                    "chart target element index must be at least -1".to_string(),
                ));
            }
            Ok(TimeVisualElement::Chart {
                shape_id_ref: read_u32(&atom.data, 8),
                build_type,
                element_index,
            })
        },
        1 => Ok(TimeVisualElement::Shape {
            kind,
            shape_id_ref: read_u32(&atom.data, 8),
            data1: read_i32(&atom.data, 12),
            data2: read_i32(&atom.data, 16),
        }),
        2 => {
            if read_u32(&atom.data, 12) != u32::MAX || read_u32(&atom.data, 16) != u32::MAX {
                return Err(PptError::InvalidFormat(
                    "VisualSoundAtom reserved data must be -1".to_string(),
                ));
            }
            Ok(TimeVisualElement::Sound {
                kind,
                sound_id_ref: read_u32(&atom.data, 8),
            })
        },
        value => Err(PptError::InvalidFormat(format!(
            "invalid visual element reference type {value}"
        ))),
    }
}

fn parse_time_variant_i32(record: &PptRecord) -> Result<i32> {
    require_time_variant_payload(record)?;
    if record.data.len() != 5 || record.data[0] != 1 {
        return Err(PptError::InvalidFormat(
            "invalid integer time variant".to_string(),
        ));
    }
    Ok(i32::from_le_bytes(
        record.data[1..5].try_into().expect("length checked"),
    ))
}

fn parse_time_variant_f32(record: &PptRecord) -> Result<f32> {
    require_time_variant_payload(record)?;
    if record.data.len() != 5 || record.data[0] != 2 {
        return Err(PptError::InvalidFormat(
            "invalid floating-point time variant".to_string(),
        ));
    }
    Ok(f32::from_le_bytes(
        record.data[1..5].try_into().expect("length checked"),
    ))
}

fn parse_time_variant_bool(record: &PptRecord) -> Result<bool> {
    require_time_variant_payload(record)?;
    if record.data.len() != 2 || record.data[0] != 0 {
        return Err(PptError::InvalidFormat(
            "invalid Boolean time variant".to_string(),
        ));
    }
    parse_bool1(record.data[1], "TimeVariant.boolValue")
}

fn parse_time_variant_string(record: &PptRecord) -> Result<String> {
    require_time_variant_payload(record)?;
    if record.data.len() % 2 != 1 || record.data.first() != Some(&3) {
        return Err(PptError::InvalidFormat(
            "invalid string time variant".to_string(),
        ));
    }
    String::from_utf16(
        &record.data[1..]
            .chunks_exact(2)
            .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
            .collect::<Vec<_>>(),
    )
    .map_err(|_| PptError::InvalidFormat("invalid UTF-16 time variant".to_string()))
}

fn require_time_variant_payload(record: &PptRecord) -> Result<()> {
    if record.data_length as usize != record.data.len() {
        return Err(PptError::Corrupted(
            "truncated TimeVariant payload".to_string(),
        ));
    }
    Ok(())
}

/// Parse an exact rotation behavior container.
pub fn parse_time_rotation_behavior(record: &PptRecord) -> Result<TimeRotationBehavior> {
    require_container(
        record,
        PptRecordType::TimeRotationBehaviorContainer,
        0,
        "TimeRotationBehaviorContainer",
    )?;
    if record.children.len() != 2 {
        return Err(PptError::InvalidFormat(
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
        return Err(PptError::InvalidFormat(
            "rotation behavior requires exactly one r or ppt_r attribute".to_string(),
        ));
    }
    validate_basic_behavior_properties(&behavior)?;
    Ok(TimeRotationBehavior { atom, behavior })
}

/// Parse an exact 20-byte `TimeRotationBehaviorAtom` payload.
pub fn parse_time_rotation_behavior_atom(record: &PptRecord) -> Result<TimeRotationBehaviorAtom> {
    require_atom(
        record,
        PptRecordType::TimeRotationBehavior,
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
        return Err(PptError::InvalidFormat(
            "rotation from value must be zero when unused".to_string(),
        ));
    };
    let to_degrees = if flags & 0x04 != 0 {
        Some(read_f32(&record.data, 12))
    } else if read_f32(&record.data, 12) == 360.0 {
        None
    } else {
        return Err(PptError::InvalidFormat(
            "rotation to value must be 360 when unused".to_string(),
        ));
    };
    let direction = if flags & 0x08 != 0 {
        Some(match read_u32(&record.data, 16) {
            0 => TimeRotationDirection::Clockwise,
            1 => TimeRotationDirection::CounterClockwise,
            value => {
                return Err(PptError::InvalidFormat(format!(
                    "invalid rotation direction {value}"
                )));
            },
        })
    } else if read_u32(&record.data, 16) == 0 {
        None
    } else {
        return Err(PptError::InvalidFormat(
            "rotation direction must be zero when unused".to_string(),
        ));
    };
    if from_degrees.is_some() && by_degrees.is_none() && to_degrees.is_none() {
        return Err(PptError::InvalidFormat(
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
pub fn parse_time_scale_behavior(record: &PptRecord) -> Result<TimeScaleBehavior> {
    require_container(
        record,
        PptRecordType::TimeScaleBehaviorContainer,
        0,
        "TimeScaleBehaviorContainer",
    )?;
    if record.children.len() != 2 {
        return Err(PptError::InvalidFormat(
            "TimeScaleBehaviorContainer requires an atom and common behavior".to_string(),
        ));
    }
    let atom = parse_time_scale_behavior_atom(&record.children[0])?;
    let behavior = parse_time_behavior(&record.children[1])?;
    validate_basic_behavior_properties(&behavior)?;
    Ok(TimeScaleBehavior { atom, behavior })
}

/// Parse an exact 32-byte `TimeScaleBehaviorAtom` payload.
pub fn parse_time_scale_behavior_atom(record: &PptRecord) -> Result<TimeScaleBehaviorAtom> {
    require_atom(
        record,
        PptRecordType::TimeScaleBehavior,
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
        return Err(PptError::InvalidFormat(
            "scale from values must be zero when unused".to_string(),
        ));
    };
    let to_percent = if flags & 0x04 != 0 {
        Some((read_f32(&record.data, 20), read_f32(&record.data, 24)))
    } else if read_f32(&record.data, 20) == 100.0 && read_f32(&record.data, 24) == 100.0 {
        None
    } else {
        return Err(PptError::InvalidFormat(
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
        return Err(PptError::InvalidFormat(
            "scale zoom-contents value must be true when unused".to_string(),
        ));
    };
    if from_percent.is_some() && by_percent.is_none() && to_percent.is_none() {
        return Err(PptError::InvalidFormat(
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
pub fn parse_time_set_behavior(record: &PptRecord) -> Result<TimeSetBehavior> {
    require_container(
        record,
        PptRecordType::TimeSetBehaviorContainer,
        0,
        "TimeSetBehaviorContainer",
    )?;
    let atom = record
        .children
        .first()
        .ok_or_else(|| PptError::InvalidFormat("set behavior has no atom".to_string()))
        .and_then(parse_time_set_behavior_atom)?;
    let mut index = 1;
    let to =
        if record.children.get(index).is_some_and(|child| {
            child.record_type == PptRecordType::TimeVariant && child.instance == 1
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
        .ok_or_else(|| PptError::InvalidFormat("set behavior has no target".to_string()))
        .and_then(parse_time_behavior)?;
    index += 1;
    if index != record.children.len() {
        return Err(PptError::InvalidFormat(
            "set behavior has invalid child order or extra children".to_string(),
        ));
    }
    let set = TimeSetBehavior { atom, to, behavior };
    validate_set_behavior(&set)?;
    Ok(set)
}

/// Parse an exact 8-byte `TimeSetBehaviorAtom` payload.
pub fn parse_time_set_behavior_atom(record: &PptRecord) -> Result<TimeSetBehaviorAtom> {
    require_atom(
        record,
        PptRecordType::TimeSetBehavior,
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
            value => {
                return Err(PptError::InvalidFormat(format!(
                    "invalid set behavior value type {value}"
                )));
            },
        })
    } else if value == 1 {
        None
    } else {
        return Err(PptError::InvalidFormat(
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
        return Err(PptError::InvalidFormat(
            "set to-use flag requires a value".to_string(),
        ));
    }
    if !set.behavior.atom.attribute_names_used {
        return Err(PptError::InvalidFormat(
            "set behavior requires an explicit attribute name".to_string(),
        ));
    }
    let attribute = match set.behavior.attribute_names.as_deref() {
        Some([attribute]) => attribute.as_str(),
        _ => {
            return Err(PptError::InvalidFormat(
                "set behavior requires exactly one attribute name".to_string(),
            ));
        },
    };
    let expected_type = time_set_attribute_value_type(attribute).ok_or_else(|| {
        PptError::InvalidFormat(format!("unsupported set behavior attribute {attribute}"))
    })?;
    let actual_type = set.atom.value_type.unwrap_or(TimeAnimateValueType::Number);
    if actual_type != expected_type {
        return Err(PptError::InvalidFormat(
            "set behavior value type does not match its attribute".to_string(),
        ));
    }
    if set
        .to
        .as_deref()
        .is_some_and(|value| !is_valid_time_set_value(attribute, value))
    {
        return Err(PptError::InvalidFormat(
            "set behavior value is invalid for its attribute".to_string(),
        ));
    }
    validate_basic_behavior_properties(&set.behavior)
}

fn validate_basic_behavior_properties(behavior: &TimeBehavior) -> Result<()> {
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
        return Err(PptError::InvalidFormat(
            "behavior contains properties reserved for color or motion behaviors".to_string(),
        ));
    }
    Ok(())
}

/// Parse an exact command behavior container.
pub fn parse_time_command_behavior(record: &PptRecord) -> Result<TimeCommandBehavior> {
    require_container(
        record,
        PptRecordType::TimeCommandBehaviorContainer,
        0,
        "TimeCommandBehaviorContainer",
    )?;
    let atom = record
        .children
        .first()
        .ok_or_else(|| PptError::InvalidFormat("command behavior has no atom".to_string()))
        .and_then(parse_time_command_behavior_atom)?;
    let mut index = 1;
    let command =
        if record.children.get(index).is_some_and(|child| {
            child.record_type == PptRecordType::TimeVariant && child.instance == 1
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
        .ok_or_else(|| PptError::InvalidFormat("command behavior has no target".to_string()))
        .and_then(parse_time_behavior)?;
    index += 1;
    if index != record.children.len() {
        return Err(PptError::InvalidFormat(
            "command behavior has invalid child order or extra children".to_string(),
        ));
    }
    validate_basic_behavior_properties(&behavior)?;
    if let Some(command) = &command {
        validate_time_command(atom.command_type, command)?;
    }
    Ok(TimeCommandBehavior {
        atom,
        command,
        behavior,
    })
}

/// Parse an exact 8-byte `TimeCommandBehaviorAtom` payload.
pub fn parse_time_command_behavior_atom(record: &PptRecord) -> Result<TimeCommandBehaviorAtom> {
    require_atom(
        record,
        PptRecordType::TimeCommandBehavior,
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
            value => {
                return Err(PptError::InvalidFormat(format!(
                    "invalid command behavior type {value}"
                )));
            },
        })
    } else if value == 1 {
        None
    } else {
        return Err(PptError::InvalidFormat(
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
        return Err(PptError::InvalidFormat(
            "invalid command for command behavior type".to_string(),
        ));
    }
    Ok(())
}

/// Parse an exact `TimeIterateDataAtom`.
pub fn parse_time_iterate_data(record: &PptRecord) -> Result<TimeIterateData> {
    require_atom(
        record,
        PptRecordType::TimeIterateData,
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
        _ => Err(PptError::InvalidFormat("invalid iteration type".into())),
    })?;
    let direction = optional_u32(flags & 1 != 0, read_u32(&record.data, 8), 1, |v| match v {
        0 => Ok(TimeIterateDirection::Backward),
        1 => Ok(TimeIterateDirection::Forward),
        _ => Err(PptError::InvalidFormat(
            "invalid iteration direction".into(),
        )),
    })?;
    let interval_type = optional_u32(flags & 8 != 0, read_u32(&record.data, 12), 0, |v| match v {
        0 => Ok(TimeIterateIntervalType::Milliseconds),
        1 => Ok(TimeIterateIntervalType::TenthsOfAPercent),
        _ => Err(PptError::InvalidFormat(
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
pub fn parse_time_sequence_data(record: &PptRecord) -> Result<TimeSequenceData> {
    require_atom(
        record,
        PptRecordType::TimeSequenceData,
        0,
        20,
        "TimeSequenceDataAtom",
    )?;
    let flags = read_u32(&record.data, 16);
    let concurrent = optional_u32(flags & 1 != 0, read_u32(&record.data, 0), 0, |v| match v {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(PptError::InvalidFormat(
            "invalid sequence concurrency".into(),
        )),
    })?;
    let next_action = optional_u32(flags & 2 != 0, read_u32(&record.data, 4), 0, |v| match v {
        0 => Ok(TimeSequenceNextAction::None),
        1 => Ok(TimeSequenceNextAction::SeekToNaturalEnd),
        _ => Err(PptError::InvalidFormat(
            "invalid next sequence action".into(),
        )),
    })?;
    let previous_action =
        optional_u32(flags & 4 != 0, read_u32(&record.data, 8), 0, |v| match v {
            0 => Ok(TimeSequencePreviousAction::None),
            1 => Ok(TimeSequencePreviousAction::SkipTimedChildren),
            _ => Err(PptError::InvalidFormat(
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
        Err(PptError::InvalidFormat(
            "unused time property has a non-default value".into(),
        ))
    }
}

/// Parse an exact `TimeConditionContainer` and its optional visual target.
pub fn parse_time_condition(record: &PptRecord) -> Result<TimeCondition> {
    require_container(
        record,
        PptRecordType::TimeConditionContainer,
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
            return Err(PptError::InvalidFormat(format!(
                "invalid time condition type {value}"
            )));
        },
    };
    let atom = record
        .children
        .first()
        .ok_or_else(|| PptError::InvalidFormat("time condition has no atom".to_string()))
        .and_then(parse_time_condition_atom)?;
    let expects_visual = atom.trigger_object == TimeTriggerObject::VisualElement;
    let visual_target = match record.children.get(1) {
        Some(target) if expects_visual => Some(parse_time_visual_element(target)?),
        Some(_) => {
            return Err(PptError::InvalidFormat(
                "only visual-element conditions can contain a visual target".to_string(),
            ));
        },
        None if expects_visual => {
            return Err(PptError::InvalidFormat(
                "visual-element condition is missing its target".to_string(),
            ));
        },
        None => None,
    };
    if record.children.len() > 2 {
        return Err(PptError::InvalidFormat(
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
pub fn parse_time_condition_atom(record: &PptRecord) -> Result<TimeConditionAtom> {
    require_atom(
        record,
        PptRecordType::TimeCondition,
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
            return Err(PptError::InvalidFormat(format!(
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
            return Err(PptError::InvalidFormat(format!(
                "invalid condition trigger event {value}"
            )));
        },
    };
    let target_id = read_u32(&record.data, 8);
    if trigger_object == TimeTriggerObject::RuntimeNodeReference && target_id != 2 {
        return Err(PptError::InvalidFormat(
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
pub fn parse_time_modifier(record: &PptRecord) -> Result<TimeModifier> {
    if record.record_type != PptRecordType::TimeModifier {
        return Err(PptError::InvalidFormat(format!(
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
        kind => Err(PptError::InvalidFormat(format!(
            "invalid time modifier type {kind}"
        ))),
    }
}

/// Parse build list from BuildList container record.
pub fn parse_build_list(record: &PptRecord) -> Result<BuildList> {
    if record.record_type != PptRecordType::BuildList {
        return Err(PptError::InvalidFormat(format!(
            "Expected BuildList record, got {:?}",
            record.record_type
        )));
    }

    require_container(record, PptRecordType::BuildList, 0, "BuildList")?;
    let mut build_info = BuildList::new();
    let mut identities = std::collections::HashSet::with_capacity(record.children.len());
    for child in &record.children {
        let build = match child.record_type {
            PptRecordType::ParaBuild => BuildListEntry::Paragraph(parse_paragraph_build(child)?),
            PptRecordType::ChartBuild => BuildListEntry::Chart(parse_chart_build(child)?),
            PptRecordType::DiagramBuild => BuildListEntry::Diagram(parse_diagram_build(child)?),
            other => {
                return Err(PptError::InvalidFormat(format!(
                    "BuildList contains invalid child {other:?}"
                )));
            },
        };
        let atom = match &build {
            BuildListEntry::Paragraph(build) => &build.atom,
            BuildListEntry::Chart(build) => &build.atom,
            BuildListEntry::Diagram(build) => &build.atom,
        };
        if !identities.insert((atom.build_id, atom.shape_id_ref)) {
            return Err(PptError::InvalidFormat(format!(
                "duplicate build identity ({}, {})",
                atom.build_id, atom.shape_id_ref
            )));
        }
        build_info.add_build(build);
    }
    Ok(build_info)
}

fn parse_build_atom(record: &PptRecord, expected: BuildKind) -> Result<BuildAtom> {
    if record.record_type != PptRecordType::BuildAtom {
        return Err(PptError::InvalidFormat(format!(
            "Expected BuildAtom, got {:?}",
            record.record_type
        )));
    }
    require_header(record, 0, 0, Some(16), "BuildAtom")?;
    let kind = BuildKind::parse(read_u32(&record.data, 0)).ok_or_else(|| {
        PptError::InvalidFormat(format!(
            "invalid BuildAtom build type {}",
            read_u32(&record.data, 0)
        ))
    })?;
    if kind != expected {
        return Err(PptError::InvalidFormat(format!(
            "BuildAtom type {kind:?} does not match {expected:?} container"
        )));
    }
    Ok(BuildAtom {
        build_id: read_u32(&record.data, 4),
        shape_id_ref: read_u32(&record.data, 8),
        expanded: parse_bool1(record.data[12], "BuildAtom.fExpanded")?,
        ui_expanded: parse_bool1(record.data[13], "BuildAtom.fUIExpanded")?,
    })
}

fn parse_paragraph_build(record: &PptRecord) -> Result<ParagraphBuild> {
    require_container(record, PptRecordType::ParaBuild, 0, "ParaBuild")?;
    if record.children.len() < 4 || (record.children.len() - 2) % 2 != 0 {
        return Err(PptError::Corrupted(
            "ParaBuild requires two atoms followed by level/time-node pairs".to_string(),
        ));
    }
    let atom = parse_build_atom(&record.children[0], BuildKind::Paragraph)?;
    let paragraph = parse_paragraph_build_atom(&record.children[1])?;
    let mut levels = Vec::with_capacity((record.children.len() - 2) / 2);
    for pair in record.children[2..].chunks_exact(2) {
        let level = parse_level_info_atom(&pair[0])?;
        let time_node = parse_extended_time_node(&pair[1])?;
        if levels
            .last()
            .is_some_and(|previous: &ParagraphBuildLevel| previous.level >= level)
        {
            return Err(PptError::InvalidFormat(
                "ParaBuild levels must be strictly increasing".to_string(),
            ));
        }
        levels.push(ParagraphBuildLevel { level, time_node });
    }
    if paragraph.build_type == ParagraphBuildType::AsAWhole && levels.len() != 1 {
        return Err(PptError::InvalidFormat(
            "AsAWhole ParaBuild requires exactly one level".to_string(),
        ));
    }
    Ok(ParagraphBuild {
        atom,
        paragraph,
        levels,
    })
}

fn parse_paragraph_build_atom(record: &PptRecord) -> Result<ParagraphBuildAtom> {
    require_atom(record, PptRecordType::ParaBuildAtom, 1, 16, "ParaBuildAtom")?;
    let build_type = ParagraphBuildType::parse(read_u32(&record.data, 0)).ok_or_else(|| {
        PptError::InvalidFormat(format!(
            "invalid ParaBuildAtom type {}",
            read_u32(&record.data, 0)
        ))
    })?;
    Ok(ParagraphBuildAtom {
        build_type,
        build_level: read_u32(&record.data, 4),
        animate_background: parse_bool1(record.data[8], "ParaBuildAtom.fAnimBackground")?,
        reverse: parse_bool1(record.data[9], "ParaBuildAtom.fReverse")?,
        user_set_animate_background: parse_bool1(
            record.data[10],
            "ParaBuildAtom.fUserSetAnimBackground",
        )?,
        automatic: parse_bool1(record.data[11], "ParaBuildAtom.fAutomatic")?,
        delay_time_ms: read_u32(&record.data, 12),
    })
}

fn parse_level_info_atom(record: &PptRecord) -> Result<u32> {
    require_atom(record, PptRecordType::LevelInfoAtom, 0, 4, "LevelInfoAtom")?;
    let level = read_u32(&record.data, 0);
    if level > 9 {
        return Err(PptError::InvalidFormat(format!(
            "LevelInfoAtom level {level} exceeds 9"
        )));
    }
    Ok(level)
}

fn parse_chart_build(record: &PptRecord) -> Result<ChartBuild> {
    require_container(record, PptRecordType::ChartBuild, 0, "ChartBuild")?;
    if record.children.len() != 2 {
        return Err(PptError::Corrupted(
            "ChartBuild requires exactly BuildAtom and ChartBuildAtom".to_string(),
        ));
    }
    let atom = parse_build_atom(&record.children[0], BuildKind::Chart)?;
    let chart_record = &record.children[1];
    require_atom(
        chart_record,
        PptRecordType::ChartBuildAtom,
        0,
        8,
        "ChartBuildAtom",
    )?;
    let build_type = ChartBuildType::parse(read_u32(&chart_record.data, 0)).ok_or_else(|| {
        PptError::InvalidFormat(format!(
            "invalid ChartBuildAtom type {}",
            read_u32(&chart_record.data, 0)
        ))
    })?;
    Ok(ChartBuild {
        atom,
        chart: ChartBuildAtom {
            build_type,
            animate_background: parse_bool1(
                chart_record.data[4],
                "ChartBuildAtom.fAnimBackground",
            )?,
        },
    })
}

fn parse_diagram_build(record: &PptRecord) -> Result<DiagramBuild> {
    require_container(record, PptRecordType::DiagramBuild, 0, "DiagramBuild")?;
    if record.children.len() != 2 {
        return Err(PptError::Corrupted(
            "DiagramBuild requires exactly BuildAtom and DiagramBuildAtom".to_string(),
        ));
    }
    let atom = parse_build_atom(&record.children[0], BuildKind::Diagram)?;
    let diagram_record = &record.children[1];
    require_atom(
        diagram_record,
        PptRecordType::DiagramBuildAtom,
        0,
        4,
        "DiagramBuildAtom",
    )?;
    let build_type =
        DiagramBuildType::parse(read_u32(&diagram_record.data, 0)).ok_or_else(|| {
            PptError::InvalidFormat(format!(
                "invalid DiagramBuildAtom type {}",
                read_u32(&diagram_record.data, 0)
            ))
        })?;
    Ok(DiagramBuild {
        atom,
        diagram: DiagramBuildAtom { build_type },
    })
}

fn require_container(
    record: &PptRecord,
    record_type: PptRecordType,
    instance: u16,
    name: &str,
) -> Result<()> {
    if record.record_type != record_type {
        return Err(PptError::InvalidFormat(format!(
            "Expected {name}, got {:?}",
            record.record_type
        )));
    }
    require_header(record, 0x0F, instance, None, name)?;
    let encoded_children_length = record.children.iter().try_fold(0usize, |length, child| {
        length.checked_add(8 + child.data.len())
    });
    if encoded_children_length != Some(record.data.len()) {
        return Err(PptError::Corrupted(format!(
            "{name} child records do not cover its complete payload"
        )));
    }
    Ok(())
}

fn require_atom(
    record: &PptRecord,
    record_type: PptRecordType,
    version: u16,
    length: usize,
    name: &str,
) -> Result<()> {
    if record.record_type != record_type {
        return Err(PptError::InvalidFormat(format!(
            "Expected {name}, got {:?}",
            record.record_type
        )));
    }
    require_header(record, version, 0, Some(length), name)
}

fn require_header(
    record: &PptRecord,
    version: u16,
    instance: u16,
    length: Option<usize>,
    name: &str,
) -> Result<()> {
    if record.version != version
        || record.instance != instance
        || record.data_length as usize != record.data.len()
        || length.is_some_and(|length| record.data.len() != length)
    {
        return Err(PptError::Corrupted(format!(
            "invalid {name} header: version {}, instance {}, length {}",
            record.version,
            record.instance,
            record.data.len()
        )));
    }
    Ok(())
}

fn parse_bool1(value: u8, field: &str) -> Result<bool> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(PptError::InvalidFormat(format!(
            "{field} has invalid bool1 value {value}"
        ))),
    }
}

fn parse_optional_time_value<T>(
    is_set: bool,
    value: u32,
    parse: impl FnOnce(u32) -> Option<T>,
    field: &str,
) -> Result<Option<T>> {
    if is_set {
        parse(value)
            .map(Some)
            .ok_or_else(|| PptError::InvalidFormat(format!("{field} has invalid value {value}")))
    } else if value == 0 {
        Ok(None)
    } else {
        Err(PptError::InvalidFormat(format!(
            "{field} must be zero when not explicitly set"
        )))
    }
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(data[offset..offset + 4].try_into().expect("length checked"))
}

fn read_i32(data: &[u8], offset: usize) -> i32 {
    i32::from_le_bytes(data[offset..offset + 4].try_into().expect("length checked"))
}

fn read_f32(data: &[u8], offset: usize) -> f32 {
    f32::from_le_bytes(data[offset..offset + 4].try_into().expect("length checked"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ppt::animation::{
        BuildInfo, ChartBuildType, DiagramBuildType, LegacyAnimationAtom, LegacyAnimationBuild,
        LegacyAnimationEffect, LegacyTextBuildSubEffect, ParagraphBuildType, write_animation_info,
        write_animation_info_atom, write_build_list, write_extended_time_node,
        write_time_animate_behavior, write_time_animate_behavior_atom,
        write_time_animation_value_atom, write_time_animation_value_list, write_time_behavior,
        write_time_behavior_atom, write_time_behavior_property_list, write_time_color_behavior,
        write_time_color_behavior_atom, write_time_command_behavior,
        write_time_command_behavior_atom, write_time_condition, write_time_condition_atom,
        write_time_effect_behavior, write_time_effect_behavior_atom, write_time_iterate_data,
        write_time_modifier, write_time_motion_behavior, write_time_motion_behavior_atom,
        write_time_node_atom, write_time_node_property_list, write_time_rotation_behavior,
        write_time_rotation_behavior_atom, write_time_scale_behavior,
        write_time_scale_behavior_atom, write_time_sequence_data, write_time_set_behavior,
        write_time_set_behavior_atom, write_time_visual_element,
    };

    fn sample_legacy_atom() -> LegacyAnimationAtom {
        LegacyAnimationAtom {
            dim_color: 0x0011_2233,
            reverse: true,
            automatic: true,
            has_sound: true,
            stop_sound: true,
            play: true,
            synchronous: true,
            hide_while_not_playing: true,
            animate_background: true,
            sound_id_ref: 42,
            delay_time_ms: 750,
            order_id: -2,
            slide_count: 3,
            build_type: LegacyAnimationBuild::Level3,
            effect: LegacyAnimationEffect::Fly,
            effect_direction: 0x1C,
            after_effect: AfterEffect::HideOnNextClick,
            text_build_sub_effect: LegacyTextBuildSubEffect::ByCharacter,
            ole_verb: 2,
        }
    }

    fn empty_time_node() -> ExtendedTimeNode {
        ExtendedTimeNode {
            atom: TimeNodeAtom::default(),
            properties: None,
            children: Vec::new(),
        }
    }

    fn sample_build_list() -> BuildList {
        BuildList {
            builds: vec![
                BuildListEntry::Paragraph(ParagraphBuild {
                    atom: BuildAtom {
                        build_id: 10,
                        shape_id_ref: 100,
                        expanded: true,
                        ui_expanded: false,
                    },
                    paragraph: ParagraphBuildAtom {
                        build_type: ParagraphBuildType::AsAWhole,
                        build_level: 4,
                        animate_background: true,
                        reverse: true,
                        user_set_animate_background: true,
                        automatic: true,
                        delay_time_ms: 750,
                    },
                    levels: vec![ParagraphBuildLevel {
                        level: 0,
                        time_node: empty_time_node(),
                    }],
                }),
                BuildListEntry::Chart(ChartBuild {
                    atom: BuildAtom {
                        build_id: 11,
                        shape_id_ref: 101,
                        expanded: false,
                        ui_expanded: true,
                    },
                    chart: ChartBuildAtom {
                        build_type: ChartBuildType::ByElementInCategory,
                        animate_background: true,
                    },
                }),
                BuildListEntry::Diagram(DiagramBuild {
                    atom: BuildAtom {
                        build_id: 12,
                        shape_id_ref: 102,
                        expanded: true,
                        ui_expanded: true,
                    },
                    diagram: DiagramBuildAtom {
                        build_type: DiagramBuildType::CounterClockwiseOut,
                    },
                }),
            ],
        }
    }

    #[test]
    fn round_trips_exact_animation_info_atoms_and_containers() {
        let atom = sample_legacy_atom();
        let bytes = write_animation_info_atom(&atom).unwrap();
        assert_eq!(bytes.len(), 36);
        let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(parse_animation_info_atom(&record).unwrap(), atom);

        let mut info = AnimationInfo::new();
        info.legacy_atom = Some(atom.clone());
        let (container, sound_ref) = write_animation_info(&info).unwrap();
        assert_eq!(sound_ref, 42);
        let (record, consumed) = PptRecord::parse(&container, 0).unwrap();
        assert_eq!(consumed, container.len());
        let parsed = parse_animation_info(&record).unwrap();
        assert_eq!(parsed.legacy_atom, Some(atom));
        assert_eq!(parsed.animation_count(), 1);
        assert_eq!(parsed.after_effect_color, Some(0x0011_2233));
        assert_eq!(parsed.iteration, IterationType::ByLetter);
    }

    #[test]
    fn rejects_malformed_animation_info_atoms() {
        let valid = write_animation_info_atom(&sample_legacy_atom()).unwrap();
        let mutations: &[(usize, u8)] = &[
            (12, 0x02), // invalid bool2 value
            (28, 0xFF), // invalid build type
            (29, 0x0F), // undefined effect
            (30, 0xFF), // invalid direction for Fly
            (31, 0x04), // invalid after effect
            (32, 0x03), // invalid text subdivision
        ];
        for &(offset, value) in mutations {
            let mut bytes = valid.clone();
            bytes[offset] = value;
            let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
            assert!(
                parse_animation_info_atom(&record).is_err(),
                "accepted mutation at byte {offset}"
            );
        }

        let mut short = valid;
        short[4..8].copy_from_slice(&27u32.to_le_bytes());
        let (record, _) = PptRecord::parse(&short, 0).unwrap();
        assert!(parse_animation_info_atom(&record).is_err());
    }

    #[test]
    fn round_trips_exact_time_node_atoms_and_envelopes() {
        let atom = TimeNodeAtom {
            fill: Some(TimeNodeFill::ResetWhenInactiveLegacy),
            restart: Some(TimeNodeRestart::NeverLegacy),
            node_type: Some(TimeNodeKind::Behavior),
            duration_ms: Some(-1),
        };
        let bytes = write_time_node_atom(&atom);
        assert_eq!(bytes.len(), 40);
        let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(parse_time_node_atom(&record).unwrap(), atom);

        let raw_child = PptRecord {
            record_type: PptRecordType::Unknown,
            record_type_raw: 0xF999,
            version: 0,
            instance: 7,
            data_length: 3,
            data: vec![1, 2, 3],
            children: Vec::new(),
        };
        let node = ExtendedTimeNode {
            atom,
            properties: Some(TimeNodePropertyList {
                properties: vec![
                    TimeNodeProperty::EffectType(TimeEffectType::Entrance),
                    TimeNodeProperty::EffectNodeType(TimeEffectNodeType::ClickEffect),
                ],
            }),
            children: vec![raw_child],
        };
        let bytes = write_extended_time_node(&node).unwrap();
        let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(parse_extended_time_node(&record).unwrap(), node);
    }

    #[test]
    fn rejects_malformed_time_node_atoms() {
        let default = write_time_node_atom(&TimeNodeAtom::default());
        let mutations: &[(usize, u32)] = &[
            (12, 1), // restart value without fRestartProperty
            (16, 1), // type value without fGroupingTypeProperty
            (20, 1), // fill value without fFillProperty
            (32, 1), // duration without fDurationProperty
        ];
        for &(offset, value) in mutations {
            let mut bytes = default.clone();
            bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
            let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
            assert!(parse_time_node_atom(&record).is_err());
        }

        let mut invalid_enum = default;
        invalid_enum[20..24].copy_from_slice(&5u32.to_le_bytes());
        invalid_enum[36..40].copy_from_slice(&1u32.to_le_bytes());
        let (record, _) = PptRecord::parse(&invalid_enum, 0).unwrap();
        assert!(parse_time_node_atom(&record).is_err());
    }

    #[test]
    fn supports_every_time_node_atom_enum_value_and_ignores_reserved_fields() {
        let fills = [
            TimeNodeFill::HoldUntilParentEnds,
            TimeNodeFill::ResetWhenInactive,
            TimeNodeFill::HoldUntilNext,
            TimeNodeFill::HoldUntilParentEndsLegacy,
            TimeNodeFill::ResetWhenInactiveLegacy,
        ];
        let restarts = [
            TimeNodeRestart::Never,
            TimeNodeRestart::Always,
            TimeNodeRestart::WhenNotActive,
            TimeNodeRestart::NeverLegacy,
        ];
        let kinds = [
            TimeNodeKind::Parallel,
            TimeNodeKind::Sequential,
            TimeNodeKind::Behavior,
            TimeNodeKind::Media,
        ];
        for fill in fills {
            for restart in restarts {
                for node_type in kinds {
                    let expected = TimeNodeAtom {
                        fill: Some(fill),
                        restart: Some(restart),
                        node_type: Some(node_type),
                        duration_ms: Some(i32::MIN),
                    };
                    let mut bytes = write_time_node_atom(&expected);
                    bytes[8..12].copy_from_slice(&u32::MAX.to_le_bytes());
                    bytes[24..28].copy_from_slice(&u32::MAX.to_le_bytes());
                    bytes[28] = 0xFF;
                    let flags = u32::from_le_bytes(bytes[36..40].try_into().unwrap()) | 0xFFFF_FFE0;
                    bytes[36..40].copy_from_slice(&flags.to_le_bytes());
                    let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
                    assert_eq!(parse_time_node_atom(&record).unwrap(), expected);
                }
            }
        }
    }

    #[test]
    fn round_trips_all_time_node_property_variants() {
        assert_eq!(PptRecordType::TimePropertyList.as_u16(), 0xF13D);
        assert_eq!(PptRecordType::TimeVariant.as_u16(), 0xF142);
        let root = TimeNodePropertyList {
            properties: vec![
                TimeNodeProperty::DisplayHidden(true),
                TimeNodeProperty::EffectId(42),
                TimeNodeProperty::EffectDirection(-7),
                TimeNodeProperty::EffectType(TimeEffectType::Entrance),
                TimeNodeProperty::AfterEffect(true),
                TimeNodeProperty::SlideCount(3),
                TimeNodeProperty::TimeFilter("0.0,0.5;1.0,1.0".to_string()),
                TimeNodeProperty::EventFilter("cancelBubble".to_string()),
                TimeNodeProperty::HideWhenStopped(false),
                TimeNodeProperty::GroupId(9),
                TimeNodeProperty::EffectNodeType(TimeEffectNodeType::InteractiveSequence),
                TimeNodeProperty::PlaceholderNode(true),
                TimeNodeProperty::MediaVolume(100_000.0),
                TimeNodeProperty::MediaMute(true),
                TimeNodeProperty::ZoomToFullScreen(false),
            ],
        };
        let bytes =
            write_time_node_property_list(&root, TimePropertyListContext::TimeNode).unwrap();
        let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(
            parse_time_node_property_list(&record, TimePropertyListContext::TimeNode).unwrap(),
            root
        );

        let subeffect = TimeNodePropertyList {
            properties: vec![
                TimeNodeProperty::DisplayHidden(false),
                TimeNodeProperty::MasterRelation(TimeMasterRelation::StartWithMaster),
                TimeNodeProperty::SubType,
                TimeNodeProperty::AfterEffect(false),
                TimeNodeProperty::PlaceholderNode(false),
                TimeNodeProperty::MediaVolume(0.0),
                TimeNodeProperty::MediaMute(false),
            ],
        };
        let bytes =
            write_time_node_property_list(&subeffect, TimePropertyListContext::SubEffect).unwrap();
        let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(
            parse_time_node_property_list(&record, TimePropertyListContext::SubEffect).unwrap(),
            subeffect
        );
    }

    #[test]
    fn rejects_invalid_time_node_property_lists() {
        let duplicate = TimeNodePropertyList {
            properties: vec![
                TimeNodeProperty::MediaMute(false),
                TimeNodeProperty::MediaMute(true),
            ],
        };
        assert!(
            write_time_node_property_list(&duplicate, TimePropertyListContext::TimeNode).is_err()
        );
        let wrong_context = TimeNodePropertyList {
            properties: vec![TimeNodeProperty::MasterRelation(
                TimeMasterRelation::DoNotStart,
            )],
        };
        assert!(
            write_time_node_property_list(&wrong_context, TimePropertyListContext::TimeNode)
                .is_err()
        );
        for invalid in [
            TimeNodeProperty::MediaVolume(f32::NAN),
            TimeNodeProperty::MediaVolume(100_001.0),
            TimeNodeProperty::TimeFilter("0.0,2.0".to_string()),
            TimeNodeProperty::EventFilter("other".to_string()),
        ] {
            let list = TimeNodePropertyList {
                properties: vec![invalid],
            };
            assert!(
                write_time_node_property_list(&list, TimePropertyListContext::TimeNode).is_err()
            );
        }

        let valid = TimeNodePropertyList {
            properties: vec![TimeNodeProperty::MediaMute(true)],
        };
        let bytes =
            write_time_node_property_list(&valid, TimePropertyListContext::TimeNode).unwrap();
        let (mut record, _) = PptRecord::parse(&bytes, 0).unwrap();
        record.children[0].data[0] = 1;
        assert!(parse_time_node_property_list(&record, TimePropertyListContext::TimeNode).is_err());
    }

    #[test]
    fn round_trips_shared_time_behaviors_and_all_properties() {
        assert_eq!(PptRecordType::TimeBehaviorContainer.as_u16(), 0xF12A);
        assert_eq!(PptRecordType::TimeBehavior.as_u16(), 0xF133);
        assert_eq!(PptRecordType::TimeClientVisualElement.as_u16(), 0xF13C);
        assert_eq!(PptRecordType::TimeVariantList.as_u16(), 0xF13E);
        let properties = TimeBehaviorPropertyList {
            properties: vec![
                TimeBehaviorProperty::UnknownPropertyList("vendor.extension".to_string()),
                TimeBehaviorProperty::RuntimeContext("GTE  PPT 12.0;PpT;".to_string()),
                TimeBehaviorProperty::MotionPathEditRelative(true),
                TimeBehaviorProperty::ColorModel(TimeColorModel::Hsl),
                TimeBehaviorProperty::ColorDirection(TimeColorDirection::CounterClockwise),
                TimeBehaviorProperty::Override,
                TimeBehaviorProperty::PathEditRotationAngle(90.0),
                TimeBehaviorProperty::PathEditRotationX(-0.5),
                TimeBehaviorProperty::PathEditRotationY(1.25),
                TimeBehaviorProperty::PointsTypes("AaFfTtSs".to_string()),
            ],
        };
        let behavior = TimeBehavior {
            atom: TimeBehaviorAtom {
                additive: Some(TimeBehaviorAdditive::Add),
                attribute_names_used: true,
            },
            attribute_names: Some(vec![
                "style.opacity".to_string(),
                "style.rotation".to_string(),
            ]),
            properties: Some(properties.clone()),
            target: TimeVisualElement::Shape {
                kind: TimeVisualElementKind::TextRange,
                shape_id_ref: 0xC03,
                data1: 0,
                data2: 12,
            },
        };

        let atom_bytes = write_time_behavior_atom(&behavior.atom);
        let (atom_record, _) = PptRecord::parse(&atom_bytes, 0).unwrap();
        assert_eq!(
            parse_time_behavior_atom(&atom_record).unwrap(),
            behavior.atom
        );

        let property_bytes = write_time_behavior_property_list(&properties).unwrap();
        let (property_record, _) = PptRecord::parse(&property_bytes, 0).unwrap();
        assert_eq!(
            parse_time_behavior_property_list(&property_record).unwrap(),
            properties
        );

        let bytes = write_time_behavior(&behavior).unwrap();
        let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(parse_time_behavior(&record).unwrap(), behavior);
    }

    #[test]
    fn round_trips_all_time_visual_element_forms() {
        let targets = [
            TimeVisualElement::Page,
            TimeVisualElement::Sound {
                kind: TimeVisualElementKind::Audio,
                sound_id_ref: 42,
            },
            TimeVisualElement::Shape {
                kind: TimeVisualElementKind::ShapeOnly,
                shape_id_ref: 100,
                data1: -7,
                data2: 9,
            },
            TimeVisualElement::Chart {
                shape_id_ref: 101,
                build_type: ChartBuildType::ByElementInSeries,
                element_index: -1,
            },
        ];
        for target in targets {
            let bytes = write_time_visual_element(&target).unwrap();
            let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
            assert_eq!(consumed, bytes.len());
            assert_eq!(parse_time_visual_element(&record).unwrap(), target);
        }
    }

    #[test]
    fn rejects_malformed_shared_time_behaviors() {
        let mut atom = write_time_behavior_atom(&TimeBehaviorAtom {
            additive: None,
            attribute_names_used: false,
        });
        atom[12..16].copy_from_slice(&1u32.to_le_bytes());
        let (atom_record, _) = PptRecord::parse(&atom, 0).unwrap();
        assert!(parse_time_behavior_atom(&atom_record).is_err());

        for property in [
            TimeBehaviorProperty::RuntimeContext("ppt 1.".to_string()),
            TimeBehaviorProperty::PointsTypes("A?".to_string()),
        ] {
            let list = TimeBehaviorPropertyList {
                properties: vec![property],
            };
            assert!(write_time_behavior_property_list(&list).is_err());
        }
        let duplicate = TimeBehaviorPropertyList {
            properties: vec![
                TimeBehaviorProperty::Override,
                TimeBehaviorProperty::Override,
            ],
        };
        assert!(write_time_behavior_property_list(&duplicate).is_err());
        let valid = TimeBehaviorPropertyList {
            properties: vec![TimeBehaviorProperty::Override],
        };
        let bytes = write_time_behavior_property_list(&valid).unwrap();
        let (mut record, _) = PptRecord::parse(&bytes, 0).unwrap();
        record.children[0].data_length += 1;
        assert!(parse_time_behavior_property_list(&record).is_err());
        assert!(
            write_time_visual_element(&TimeVisualElement::Shape {
                kind: TimeVisualElementKind::ChartElement,
                shape_id_ref: 1,
                data1: 0,
                data2: 0,
            })
            .is_err()
        );
        assert!(
            write_time_visual_element(&TimeVisualElement::Chart {
                shape_id_ref: 1,
                build_type: ChartBuildType::AsOneObject,
                element_index: -2,
            })
            .is_err()
        );

        let sound = TimeVisualElement::Sound {
            kind: TimeVisualElementKind::Audio,
            sound_id_ref: 42,
        };
        let bytes = write_time_visual_element(&sound).unwrap();
        let (mut record, _) = PptRecord::parse(&bytes, 0).unwrap();
        record.children[0].data[12..16].copy_from_slice(&0u32.to_le_bytes());
        assert!(parse_time_visual_element(&record).is_err());
    }

    #[test]
    fn round_trips_color_behaviors_and_color_models() {
        assert_eq!(PptRecordType::TimeColorBehaviorContainer.as_u16(), 0xF12C);
        assert_eq!(PptRecordType::TimeColorBehavior.as_u16(), 0xF135);

        for by in [
            TimeAnimateColorBy::Rgb {
                red: -255,
                green: 0,
                blue: 255,
            },
            TimeAnimateColorBy::Hsl {
                hue: 120,
                saturation: -40,
                luminance: 15,
            },
            TimeAnimateColorBy::Scheme(7),
        ] {
            let expected = TimeColorBehaviorAtom {
                by: Some(by),
                from: Some(TimeAnimateColor::Rgb {
                    red: 1,
                    green: 2,
                    blue: 255,
                }),
                to: Some(TimeAnimateColor::Scheme(3)),
                color_space_used: true,
                direction_used: true,
            };
            let bytes = write_time_color_behavior_atom(&expected).unwrap();
            let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
            assert_eq!(consumed, bytes.len());
            assert_eq!(parse_time_color_behavior_atom(&record).unwrap(), expected);
        }

        let expected = TimeColorBehavior {
            atom: TimeColorBehaviorAtom {
                by: Some(TimeAnimateColorBy::Hsl {
                    hue: 45,
                    saturation: 20,
                    luminance: -10,
                }),
                from: None,
                to: Some(TimeAnimateColor::Rgb {
                    red: 0x11,
                    green: 0x22,
                    blue: 0x33,
                }),
                color_space_used: true,
                direction_used: true,
            },
            behavior: TimeBehavior {
                atom: TimeBehaviorAtom {
                    additive: Some(TimeBehaviorAdditive::Override),
                    attribute_names_used: true,
                },
                attribute_names: Some(vec!["fill.color".to_string()]),
                properties: Some(TimeBehaviorPropertyList {
                    properties: vec![
                        TimeBehaviorProperty::RuntimeContext("ppt".to_string()),
                        TimeBehaviorProperty::ColorModel(TimeColorModel::Hsl),
                        TimeBehaviorProperty::ColorDirection(TimeColorDirection::CounterClockwise),
                    ],
                }),
                target: TimeVisualElement::Shape {
                    kind: TimeVisualElementKind::Shape,
                    shape_id_ref: 17,
                    data1: 0,
                    data2: 0,
                },
            },
        };
        let bytes = write_time_color_behavior(&expected).unwrap();
        let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(parse_time_color_behavior(&record).unwrap(), expected);
    }

    #[test]
    fn rejects_malformed_color_behaviors() {
        for atom in [
            TimeColorBehaviorAtom {
                by: Some(TimeAnimateColorBy::Rgb {
                    red: 256,
                    green: 0,
                    blue: 0,
                }),
                from: None,
                to: None,
                color_space_used: false,
                direction_used: false,
            },
            TimeColorBehaviorAtom {
                by: None,
                from: Some(TimeAnimateColor::Rgb {
                    red: 0,
                    green: 0,
                    blue: 0,
                }),
                to: None,
                color_space_used: false,
                direction_used: false,
            },
            TimeColorBehaviorAtom {
                by: Some(TimeAnimateColorBy::Scheme(8)),
                from: None,
                to: None,
                color_space_used: false,
                direction_used: false,
            },
            TimeColorBehaviorAtom {
                by: None,
                from: None,
                to: Some(TimeAnimateColor::Rgb {
                    red: 256,
                    green: 0,
                    blue: 0,
                }),
                color_space_used: false,
                direction_used: false,
            },
        ] {
            assert!(write_time_color_behavior_atom(&atom).is_err());
        }

        let valid_atom = TimeColorBehaviorAtom {
            by: Some(TimeAnimateColorBy::Rgb {
                red: 1,
                green: 2,
                blue: 3,
            }),
            from: None,
            to: None,
            color_space_used: false,
            direction_used: false,
        };
        let mut bytes = write_time_color_behavior_atom(&valid_atom).unwrap();
        bytes[12..16].copy_from_slice(&3u32.to_le_bytes());
        let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
        assert!(parse_time_color_behavior_atom(&record).is_err());

        let common = |name: &str, properties: Vec<TimeBehaviorProperty>| TimeBehavior {
            atom: TimeBehaviorAtom {
                additive: None,
                attribute_names_used: true,
            },
            attribute_names: Some(vec![name.to_string()]),
            properties: Some(TimeBehaviorPropertyList { properties }),
            target: TimeVisualElement::Page,
        };
        for invalid in [
            TimeColorBehavior {
                atom: TimeColorBehaviorAtom {
                    color_space_used: true,
                    ..valid_atom.clone()
                },
                behavior: common("fill.color", vec![]),
            },
            TimeColorBehavior {
                atom: valid_atom.clone(),
                behavior: common("style.opacity", vec![]),
            },
            TimeColorBehavior {
                atom: valid_atom,
                behavior: common(
                    "fill.color",
                    vec![TimeBehaviorProperty::MotionPathEditRelative(true)],
                ),
            },
        ] {
            assert!(write_time_color_behavior(&invalid).is_err());
        }
    }

    #[test]
    fn round_trips_all_image_effect_filters() {
        assert_eq!(PptRecordType::TimeEffectBehaviorContainer.as_u16(), 0xF12D);
        assert_eq!(PptRecordType::TimeEffectBehavior.as_u16(), 0xF136);
        let filters = [
            TimeEffectFilter::BlindsHorizontal,
            TimeEffectFilter::BlindsVertical,
            TimeEffectFilter::BoxIn,
            TimeEffectFilter::BoxOut,
            TimeEffectFilter::CheckerboardAcross,
            TimeEffectFilter::CheckerboardDown,
            TimeEffectFilter::CircleIn,
            TimeEffectFilter::CircleOut,
            TimeEffectFilter::DiamondIn,
            TimeEffectFilter::DiamondOut,
            TimeEffectFilter::Dissolve,
            TimeEffectFilter::Fade,
            TimeEffectFilter::PlusIn,
            TimeEffectFilter::PlusOut,
            TimeEffectFilter::BarnInVertical,
            TimeEffectFilter::BarnInHorizontal,
            TimeEffectFilter::BarnOutVertical,
            TimeEffectFilter::BarnOutHorizontal,
            TimeEffectFilter::RandomBarHorizontal,
            TimeEffectFilter::RandomBarVertical,
            TimeEffectFilter::StripsDownLeft,
            TimeEffectFilter::StripsUpLeft,
            TimeEffectFilter::StripsDownRight,
            TimeEffectFilter::StripsUpRight,
            TimeEffectFilter::Wedge,
            TimeEffectFilter::Wheel1,
            TimeEffectFilter::Wheel2,
            TimeEffectFilter::Wheel3,
            TimeEffectFilter::Wheel4,
            TimeEffectFilter::Wheel8,
            TimeEffectFilter::WipeRight,
            TimeEffectFilter::WipeLeft,
            TimeEffectFilter::WipeUp,
            TimeEffectFilter::WipeDown,
        ];
        let common = || TimeBehavior {
            atom: TimeBehaviorAtom {
                additive: Some(TimeBehaviorAdditive::Override),
                attribute_names_used: false,
            },
            attribute_names: Some(vec!["ignored".to_string()]),
            properties: Some(TimeBehaviorPropertyList {
                properties: vec![TimeBehaviorProperty::RuntimeContext("ppt".to_string())],
            }),
            target: TimeVisualElement::Shape {
                kind: TimeVisualElementKind::Shape,
                shape_id_ref: 21,
                data1: 0,
                data2: 0,
            },
        };
        for filter in filters {
            assert_eq!(TimeEffectFilter::parse(filter.as_str()), Some(filter));
            let expected = TimeEffectBehavior {
                atom: TimeEffectBehaviorAtom {
                    transition: Some(TimeEffectTransition::Out),
                    filter_used: true,
                    progress_used: true,
                    runtime_context_used: true,
                },
                filter: Some(filter),
                progress: Some(0.625),
                runtime_context: Some("GTE PPT 10.0;PpT;".to_string()),
                behavior: common(),
            };
            let bytes = write_time_effect_behavior(&expected).unwrap();
            let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
            assert_eq!(consumed, bytes.len());
            assert_eq!(parse_time_effect_behavior(&record).unwrap(), expected);
        }

        for transition in [
            None,
            Some(TimeEffectTransition::In),
            Some(TimeEffectTransition::Out),
        ] {
            let expected = TimeEffectBehaviorAtom {
                transition,
                filter_used: false,
                progress_used: false,
                runtime_context_used: false,
            };
            let bytes = write_time_effect_behavior_atom(&expected);
            let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
            assert_eq!(parse_time_effect_behavior_atom(&record).unwrap(), expected);
        }
    }

    #[test]
    fn rejects_malformed_image_effect_behaviors() {
        let common = || TimeBehavior {
            atom: TimeBehaviorAtom {
                additive: None,
                attribute_names_used: false,
            },
            attribute_names: None,
            properties: None,
            target: TimeVisualElement::Page,
        };
        let valid = TimeEffectBehavior {
            atom: TimeEffectBehaviorAtom {
                transition: None,
                filter_used: true,
                progress_used: true,
                runtime_context_used: true,
            },
            filter: Some(TimeEffectFilter::Fade),
            progress: Some(0.5),
            runtime_context: Some("ppt".to_string()),
            behavior: common(),
        };
        for invalid in [
            TimeEffectBehavior {
                filter: None,
                ..valid.clone()
            },
            TimeEffectBehavior {
                progress: None,
                ..valid.clone()
            },
            TimeEffectBehavior {
                runtime_context: None,
                ..valid.clone()
            },
            TimeEffectBehavior {
                progress: Some(-0.01),
                ..valid.clone()
            },
            TimeEffectBehavior {
                progress: Some(f32::NAN),
                ..valid.clone()
            },
            TimeEffectBehavior {
                runtime_context: Some("ppt 1.".to_string()),
                ..valid.clone()
            },
            TimeEffectBehavior {
                behavior: TimeBehavior {
                    properties: Some(TimeBehaviorPropertyList {
                        properties: vec![TimeBehaviorProperty::ColorModel(TimeColorModel::Rgb)],
                    }),
                    ..common()
                },
                ..valid.clone()
            },
        ] {
            assert!(write_time_effect_behavior(&invalid).is_err());
        }

        let mut bytes = write_time_effect_behavior_atom(&TimeEffectBehaviorAtom {
            transition: None,
            filter_used: false,
            progress_used: false,
            runtime_context_used: false,
        });
        bytes[12..16].copy_from_slice(&1u32.to_le_bytes());
        let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
        assert!(parse_time_effect_behavior_atom(&record).is_err());

        let bytes = write_time_effect_behavior(&valid).unwrap();
        let (mut record, _) = PptRecord::parse(&bytes, 0).unwrap();
        record.children[1].data = vec![3, b'n', 0, b'o', 0, b'p', 0, b'e', 0];
        record.children[1].data_length = 9;
        assert!(parse_time_effect_behavior(&record).is_err());

        let (mut record, _) = PptRecord::parse(&bytes, 0).unwrap();
        record.children.swap(1, 2);
        assert!(parse_time_effect_behavior(&record).is_err());
    }

    #[test]
    fn round_trips_motion_behaviors_and_formula_paths() {
        assert_eq!(PptRecordType::TimeMotionBehaviorContainer.as_u16(), 0xF12E);
        assert_eq!(PptRecordType::TimeMotionBehavior.as_u16(), 0xF137);
        let path = "M 0 0 L 1.0 (ppt_x+$) C 0 0.5 (sin(pi)) 1 1 (max(#ppt_y,0.25)) Z E ignored";
        let expected = TimeMotionBehavior {
            atom: TimeMotionBehaviorAtom {
                by: Some((0.25, -0.5)),
                from: Some((0.0, 0.0)),
                to: Some((1.0, 1.0)),
                origin: Some(TimeMotionOrigin::ObjectCenter),
                path_used: true,
                edit_rotation_used: true,
                points_types_used: true,
            },
            path: Some(path.to_string()),
            reserved: Some(-7),
            behavior: TimeBehavior {
                atom: TimeBehaviorAtom {
                    additive: Some(TimeBehaviorAdditive::Add),
                    attribute_names_used: true,
                },
                attribute_names: Some(vec!["ppt_x".to_string(), "ppt_y".to_string()]),
                properties: Some(TimeBehaviorPropertyList {
                    properties: vec![
                        TimeBehaviorProperty::MotionPathEditRelative(true),
                        TimeBehaviorProperty::PathEditRotationAngle(45.0),
                        TimeBehaviorProperty::PathEditRotationX(0.5),
                        TimeBehaviorProperty::PathEditRotationY(0.5),
                        TimeBehaviorProperty::PointsTypes("AaFfTtSs".to_string()),
                    ],
                }),
                target: TimeVisualElement::Shape {
                    kind: TimeVisualElementKind::Shape,
                    shape_id_ref: 22,
                    data1: 0,
                    data2: 0,
                },
            },
        };
        let bytes = write_time_motion_behavior(&expected).unwrap();
        let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(parse_time_motion_behavior(&record).unwrap(), expected);

        for origin in [
            None,
            Some(TimeMotionOrigin::Slide),
            Some(TimeMotionOrigin::SlideLegacy),
            Some(TimeMotionOrigin::ObjectCenter),
        ] {
            let expected = TimeMotionBehaviorAtom {
                by: None,
                from: None,
                to: None,
                origin,
                path_used: false,
                edit_rotation_used: false,
                points_types_used: false,
            };
            let bytes = write_time_motion_behavior_atom(&expected).unwrap();
            let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
            assert_eq!(parse_time_motion_behavior_atom(&record).unwrap(), expected);
        }
    }

    #[test]
    fn rejects_malformed_motion_behaviors_and_paths() {
        let common = || TimeBehavior {
            atom: TimeBehaviorAtom {
                additive: None,
                attribute_names_used: false,
            },
            attribute_names: None,
            properties: None,
            target: TimeVisualElement::Page,
        };
        let atom = TimeMotionBehaviorAtom {
            by: Some((1.0, 1.0)),
            from: None,
            to: None,
            origin: None,
            path_used: true,
            edit_rotation_used: false,
            points_types_used: false,
        };
        for path in [
            "",
            "Q 0 0",
            "M -1 0",
            "M .5 0",
            "M 1. 0",
            "M (unknown) 0",
            "M (max(1,2,3)) 0",
            "M (sin( 1)) 0",
            "C 0 0 1 1",
            "M 0 0 X",
        ] {
            let invalid = TimeMotionBehavior {
                atom: atom.clone(),
                path: Some(path.to_string()),
                reserved: None,
                behavior: common(),
            };
            assert!(write_time_motion_behavior(&invalid).is_err(), "{path}");
        }

        let mut invalid_atom = atom.clone();
        invalid_atom.by = None;
        invalid_atom.from = Some((0.0, 0.0));
        assert!(write_time_motion_behavior_atom(&invalid_atom).is_err());
        let mut bytes = write_time_motion_behavior_atom(&TimeMotionBehaviorAtom {
            path_used: false,
            ..atom.clone()
        })
        .unwrap();
        bytes[36..40].copy_from_slice(&0u32.to_le_bytes());
        let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
        assert!(parse_time_motion_behavior_atom(&record).is_err());

        for invalid in [
            TimeMotionBehavior {
                atom: atom.clone(),
                path: None,
                reserved: None,
                behavior: common(),
            },
            TimeMotionBehavior {
                atom: TimeMotionBehaviorAtom {
                    edit_rotation_used: true,
                    ..atom.clone()
                },
                path: Some("M 0 0".to_string()),
                reserved: None,
                behavior: common(),
            },
            TimeMotionBehavior {
                atom: TimeMotionBehaviorAtom {
                    points_types_used: true,
                    ..atom.clone()
                },
                path: Some("M 0 0".to_string()),
                reserved: None,
                behavior: common(),
            },
            TimeMotionBehavior {
                atom: atom.clone(),
                path: Some("M 0 0".to_string()),
                reserved: None,
                behavior: TimeBehavior {
                    properties: Some(TimeBehaviorPropertyList {
                        properties: vec![TimeBehaviorProperty::ColorDirection(
                            TimeColorDirection::Clockwise,
                        )],
                    }),
                    ..common()
                },
            },
            TimeMotionBehavior {
                atom,
                path: Some("M 0 0".to_string()),
                reserved: None,
                behavior: TimeBehavior {
                    atom: TimeBehaviorAtom {
                        additive: None,
                        attribute_names_used: true,
                    },
                    attribute_names: Some(vec!["ppt_x".into(), "ppt_y".into(), "ppt_w".into()]),
                    ..common()
                },
            },
        ] {
            assert!(write_time_motion_behavior(&invalid).is_err());
        }
    }

    #[test]
    fn round_trips_rotation_and_scale_behaviors() {
        assert_eq!(
            PptRecordType::TimeRotationBehaviorContainer.as_u16(),
            0xF12F
        );
        assert_eq!(PptRecordType::TimeScaleBehaviorContainer.as_u16(), 0xF130);
        assert_eq!(PptRecordType::TimeRotationBehavior.as_u16(), 0xF138);
        assert_eq!(PptRecordType::TimeScaleBehavior.as_u16(), 0xF139);
        let common = |attribute_names: Option<Vec<String>>, used| TimeBehavior {
            atom: TimeBehaviorAtom {
                additive: Some(TimeBehaviorAdditive::Override),
                attribute_names_used: used,
            },
            attribute_names,
            properties: Some(TimeBehaviorPropertyList {
                properties: vec![TimeBehaviorProperty::RuntimeContext("ppt 12".to_string())],
            }),
            target: TimeVisualElement::Shape {
                kind: TimeVisualElementKind::Shape,
                shape_id_ref: 7,
                data1: 0,
                data2: 0,
            },
        };
        let rotation = TimeRotationBehavior {
            atom: TimeRotationBehaviorAtom {
                by_degrees: Some(45.0),
                from_degrees: Some(-15.0),
                to_degrees: Some(180.0),
                direction: Some(TimeRotationDirection::CounterClockwise),
            },
            behavior: common(Some(vec!["ppt_r".to_string()]), true),
        };
        let atom_bytes = write_time_rotation_behavior_atom(&rotation.atom).unwrap();
        let (atom_record, _) = PptRecord::parse(&atom_bytes, 0).unwrap();
        assert_eq!(
            parse_time_rotation_behavior_atom(&atom_record).unwrap(),
            rotation.atom
        );
        let bytes = write_time_rotation_behavior(&rotation).unwrap();
        let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(parse_time_rotation_behavior(&record).unwrap(), rotation);

        let scale = TimeScaleBehavior {
            atom: TimeScaleBehaviorAtom {
                by_percent: Some((10.0, 20.0)),
                from_percent: Some((80.0, 90.0)),
                to_percent: Some((120.0, 130.0)),
                zoom_contents: Some(false),
            },
            behavior: common(Some(vec!["ignored".to_string()]), false),
        };
        let atom_bytes = write_time_scale_behavior_atom(&scale.atom).unwrap();
        let (atom_record, _) = PptRecord::parse(&atom_bytes, 0).unwrap();
        assert_eq!(
            parse_time_scale_behavior_atom(&atom_record).unwrap(),
            scale.atom
        );
        let bytes = write_time_scale_behavior(&scale).unwrap();
        let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(parse_time_scale_behavior(&record).unwrap(), scale);
    }

    #[test]
    fn rejects_malformed_rotation_and_scale_behaviors() {
        let invalid_rotation = TimeRotationBehaviorAtom {
            by_degrees: None,
            from_degrees: Some(1.0),
            to_degrees: None,
            direction: None,
        };
        assert!(write_time_rotation_behavior_atom(&invalid_rotation).is_err());
        let invalid_scale = TimeScaleBehaviorAtom {
            by_percent: None,
            from_percent: Some((1.0, 1.0)),
            to_percent: None,
            zoom_contents: None,
        };
        assert!(write_time_scale_behavior_atom(&invalid_scale).is_err());

        let mut bytes = write_time_rotation_behavior_atom(&TimeRotationBehaviorAtom {
            by_degrees: None,
            from_degrees: None,
            to_degrees: None,
            direction: None,
        })
        .unwrap();
        bytes[20..24].copy_from_slice(&0f32.to_le_bytes());
        let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
        assert!(parse_time_rotation_behavior_atom(&record).is_err());

        let mut bytes = write_time_scale_behavior_atom(&TimeScaleBehaviorAtom {
            by_percent: None,
            from_percent: None,
            to_percent: None,
            zoom_contents: None,
        })
        .unwrap();
        bytes[36] = 0;
        let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
        assert!(parse_time_scale_behavior_atom(&record).is_err());
    }

    #[test]
    fn round_trips_generic_animate_behaviors_and_keyframes() {
        assert_eq!(PptRecordType::TimeAnimateBehaviorContainer.as_u16(), 0xF12B);
        assert_eq!(PptRecordType::TimeAnimateBehavior.as_u16(), 0xF134);
        assert_eq!(PptRecordType::TimeAnimationValueList.as_u16(), 0xF13F);
        assert_eq!(PptRecordType::TimeAnimationValue.as_u16(), 0xF143);
        let common = |attribute: &str| TimeBehavior {
            atom: TimeBehaviorAtom {
                additive: Some(TimeBehaviorAdditive::Override),
                attribute_names_used: true,
            },
            attribute_names: Some(vec![attribute.to_string()]),
            properties: Some(TimeBehaviorPropertyList {
                properties: vec![TimeBehaviorProperty::RuntimeContext("ppt".to_string())],
            }),
            target: TimeVisualElement::Shape {
                kind: TimeVisualElementKind::Shape,
                shape_id_ref: 24,
                data1: 0,
                data2: 0,
            },
        };
        let values = TimeAnimationValueList {
            entries: vec![
                TimeAnimationValue {
                    time: -1000,
                    value: Some(TimeVariantValue::Boolean(true)),
                    formula: None,
                },
                TimeAnimationValue {
                    time: 333,
                    value: Some(TimeVariantValue::Integer(-2)),
                    formula: Some("max($,#ppt_y)".to_string()),
                },
                TimeAnimationValue {
                    time: 667,
                    value: Some(TimeVariantValue::Float(1.25)),
                    formula: None,
                },
                TimeAnimationValue {
                    time: 1000,
                    value: Some(TimeVariantValue::String("2".to_string())),
                    formula: None,
                },
            ],
        };
        let expected = TimeAnimateBehavior {
            atom: TimeAnimateBehaviorAtom {
                calculation_mode: Some(TimeAnimateCalculationMode::Formula),
                by_used: true,
                from_used: true,
                to_used: true,
                animation_values_used: true,
                value_type: None,
            },
            values: Some(values.clone()),
            by: Some("1".to_string()),
            from: Some("0".to_string()),
            to: Some("2".to_string()),
            behavior: common("ppt_x"),
        };
        let bytes = write_time_animate_behavior(&expected).unwrap();
        let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(parse_time_animate_behavior(&record).unwrap(), expected);

        let bytes = write_time_animation_value_list(&values).unwrap();
        let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(parse_time_animation_value_list(&record).unwrap(), values);

        for (attribute, value_type, value) in [
            ("image", TimeAnimateValueType::String, "arbitrary 👋"),
            ("fill.color", TimeAnimateValueType::Color, "#A0b1C2"),
        ] {
            let expected = TimeAnimateBehavior {
                atom: TimeAnimateBehaviorAtom {
                    calculation_mode: Some(TimeAnimateCalculationMode::Discrete),
                    by_used: true,
                    from_used: false,
                    to_used: true,
                    animation_values_used: false,
                    value_type: Some(value_type),
                },
                values: None,
                by: Some(value.to_string()),
                from: None,
                to: Some(value.to_string()),
                behavior: common(attribute),
            };
            let bytes = write_time_animate_behavior(&expected).unwrap();
            let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
            assert_eq!(parse_time_animate_behavior(&record).unwrap(), expected);
        }

        for mode in [
            None,
            Some(TimeAnimateCalculationMode::Discrete),
            Some(TimeAnimateCalculationMode::Linear),
            Some(TimeAnimateCalculationMode::Formula),
        ] {
            for value_type in [
                None,
                Some(TimeAnimateValueType::String),
                Some(TimeAnimateValueType::Number),
                Some(TimeAnimateValueType::Color),
            ] {
                let expected = TimeAnimateBehaviorAtom {
                    calculation_mode: mode,
                    by_used: false,
                    from_used: false,
                    to_used: false,
                    animation_values_used: false,
                    value_type,
                };
                let bytes = write_time_animate_behavior_atom(&expected);
                let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
                assert_eq!(parse_time_animate_behavior_atom(&record).unwrap(), expected);
            }
        }
    }

    #[test]
    fn rejects_malformed_generic_animate_behaviors() {
        let common = |attribute: &str| TimeBehavior {
            atom: TimeBehaviorAtom {
                additive: None,
                attribute_names_used: true,
            },
            attribute_names: Some(vec![attribute.to_string()]),
            properties: None,
            target: TimeVisualElement::Page,
        };
        let valid = TimeAnimateBehavior {
            atom: TimeAnimateBehaviorAtom {
                calculation_mode: None,
                by_used: true,
                from_used: false,
                to_used: false,
                animation_values_used: false,
                value_type: None,
            },
            values: None,
            by: Some("1".to_string()),
            from: None,
            to: None,
            behavior: common("ppt_x"),
        };
        for invalid in [
            TimeAnimateBehavior {
                by: None,
                ..valid.clone()
            },
            TimeAnimateBehavior {
                atom: TimeAnimateBehaviorAtom {
                    animation_values_used: true,
                    ..valid.atom.clone()
                },
                ..valid.clone()
            },
            TimeAnimateBehavior {
                atom: TimeAnimateBehaviorAtom {
                    by_used: false,
                    ..valid.atom.clone()
                },
                by: None,
                from: Some("0".to_string()),
                ..valid.clone()
            },
            TimeAnimateBehavior {
                atom: TimeAnimateBehaviorAtom {
                    value_type: Some(TimeAnimateValueType::Color),
                    ..valid.atom.clone()
                },
                ..valid.clone()
            },
            TimeAnimateBehavior {
                by: Some("invalid".to_string()),
                ..valid.clone()
            },
            TimeAnimateBehavior {
                behavior: common("unsupported.attribute"),
                ..valid.clone()
            },
            TimeAnimateBehavior {
                atom: TimeAnimateBehaviorAtom {
                    calculation_mode: Some(TimeAnimateCalculationMode::Formula),
                    ..valid.atom.clone()
                },
                ..valid.clone()
            },
        ] {
            assert!(write_time_animate_behavior(&invalid).is_err());
        }

        for time in [-1001, 1001] {
            assert!(write_time_animation_value_atom(time).is_err());
        }
        let invalid_list = TimeAnimationValueList {
            entries: vec![TimeAnimationValue {
                time: 0,
                value: None,
                formula: Some("unknown+1".to_string()),
            }],
        };
        assert!(write_time_animation_value_list(&invalid_list).is_err());

        let mut atom = write_time_animate_behavior_atom(&TimeAnimateBehaviorAtom {
            calculation_mode: None,
            by_used: false,
            from_used: false,
            to_used: false,
            animation_values_used: false,
            value_type: None,
        });
        atom[8..12].copy_from_slice(&0u32.to_le_bytes());
        let (record, _) = PptRecord::parse(&atom, 0).unwrap();
        assert!(parse_time_animate_behavior_atom(&record).is_err());
    }

    #[test]
    fn round_trips_set_behaviors_for_all_value_categories() {
        assert_eq!(PptRecordType::TimeSetBehaviorContainer.as_u16(), 0xF131);
        assert_eq!(PptRecordType::TimeSetBehavior.as_u16(), 0xF13A);
        let common = |attribute: &str| TimeBehavior {
            atom: TimeBehaviorAtom {
                additive: Some(TimeBehaviorAdditive::Override),
                attribute_names_used: true,
            },
            attribute_names: Some(vec![attribute.to_string()]),
            properties: Some(TimeBehaviorPropertyList {
                properties: vec![TimeBehaviorProperty::RuntimeContext("ppt".to_string())],
            }),
            target: TimeVisualElement::Shape {
                kind: TimeVisualElementKind::Shape,
                shape_id_ref: 23,
                data1: 0,
                data2: 0,
            },
        };
        let cases = [
            ("style.visibility", TimeAnimateValueType::Number, "hidden"),
            ("style.fontWeight", TimeAnimateValueType::Number, "bold"),
            ("fill.type", TimeAnimateValueType::Number, "gradientRadial"),
            (
                "stroke.dashstyle",
                TimeAnimateValueType::Number,
                "longDashDotDot",
            ),
            (
                "stroke.startArrow",
                TimeAnimateValueType::Number,
                "doublechevron",
            ),
            (
                "extrusion.render",
                TimeAnimateValueType::Number,
                "boundingcube",
            ),
            (
                "ppt_x",
                TimeAnimateValueType::Number,
                "(max($,#ppt_y)+1.5e2)",
            ),
            ("shadow.matrix.ytoy", TimeAnimateValueType::Number, "-.5"),
            (
                "extrusion.rotationcenter.z",
                TimeAnimateValueType::Number,
                "1-e2",
            ),
            ("ppt_c", TimeAnimateValueType::Color, "#00aF7C"),
            ("extrusion.color", TimeAnimateValueType::Color, "#AABBCC"),
        ];
        for (attribute, value_type, value) in cases {
            let expected = TimeSetBehavior {
                atom: TimeSetBehaviorAtom {
                    to_used: true,
                    value_type: (value_type != TimeAnimateValueType::Number).then_some(value_type),
                },
                to: Some(value.to_string()),
                behavior: common(attribute),
            };
            let bytes = write_time_set_behavior(&expected).unwrap();
            let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
            assert_eq!(consumed, bytes.len());
            assert_eq!(parse_time_set_behavior(&record).unwrap(), expected);
        }

        for value_type in [
            None,
            Some(TimeAnimateValueType::String),
            Some(TimeAnimateValueType::Number),
            Some(TimeAnimateValueType::Color),
        ] {
            let expected = TimeSetBehaviorAtom {
                to_used: false,
                value_type,
            };
            let bytes = write_time_set_behavior_atom(&expected);
            let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
            assert_eq!(parse_time_set_behavior_atom(&record).unwrap(), expected);
        }
    }

    #[test]
    fn validates_set_presets_numbers_formulas_and_colors() {
        let presets = [
            ("style.visibility", "visible"),
            ("style.fontStyle", "italic"),
            ("style.textEffectEmboss", "emboss"),
            ("style.textShadow", "auto"),
            ("style.textTransform", "super"),
            ("style.textDecorationUnderline", "true"),
            ("style.textEffectOutline", "false"),
            ("style.textDecorationLineThrough", "true"),
            ("imageData.grayscale", "false"),
            ("fill.on", "t"),
            ("fill.method", "sigma"),
            ("stroke.on", "f"),
            ("stroke.linestyle", "thickBetweenThin"),
            ("stroke.filltype", "frame"),
            ("stroke.endArrow", "chevron"),
            ("stroke.startArrowWidth", "narrow"),
            ("stroke.startArrowLength", "long"),
            ("stroke.endArrowWidth", "wide"),
            ("stroke.endArrowLength", "short"),
            ("shadow.on", "true"),
            ("shadow.type", "perspective"),
            ("skew.on", "false"),
            ("extrusion.on", "true"),
            ("extrusion.type", "parallel"),
            ("extrusion.plane", "yz"),
            ("extrusion.lockrotationcenter", "false"),
            ("extrusion.autorotationcenter", "true"),
            ("extrusion.colormode", "false"),
        ];
        for (attribute, value) in presets {
            assert_eq!(
                time_set_attribute_value_type(attribute),
                Some(TimeAnimateValueType::Number)
            );
            assert!(is_valid_time_set_value(attribute, value));
        }
        for value in ["0", "-1", "1.", ".5", "-.5", "1e2", "1-e2", "(sqrt(4))"] {
            assert!(is_valid_time_set_value("ppt_x", value), "{value}");
        }
        for attribute in [
            "ppt_c",
            "fillcolor",
            "style.color",
            "imageData.chromakey",
            "fill.color",
            "fill.color2",
            "stroke.color",
            "stroke.color2",
            "shadow.color",
            "shadow.color2",
            "extrusion.color",
        ] {
            assert_eq!(
                time_set_attribute_value_type(attribute),
                Some(TimeAnimateValueType::Color)
            );
            assert!(is_valid_time_set_value(attribute, "#123abc"));
        }
    }

    #[test]
    fn rejects_malformed_set_behaviors() {
        let common = |attribute_names: Option<Vec<String>>, used| TimeBehavior {
            atom: TimeBehaviorAtom {
                additive: None,
                attribute_names_used: used,
            },
            attribute_names,
            properties: None,
            target: TimeVisualElement::Page,
        };
        let valid = TimeSetBehavior {
            atom: TimeSetBehaviorAtom {
                to_used: true,
                value_type: None,
            },
            to: Some("visible".to_string()),
            behavior: common(Some(vec!["style.visibility".to_string()]), true),
        };
        for invalid in [
            TimeSetBehavior {
                to: None,
                ..valid.clone()
            },
            TimeSetBehavior {
                to: Some("opaque".to_string()),
                ..valid.clone()
            },
            TimeSetBehavior {
                atom: TimeSetBehaviorAtom {
                    to_used: true,
                    value_type: Some(TimeAnimateValueType::Color),
                },
                ..valid.clone()
            },
            TimeSetBehavior {
                atom: TimeSetBehaviorAtom {
                    to_used: true,
                    value_type: Some(TimeAnimateValueType::String),
                },
                ..valid.clone()
            },
            TimeSetBehavior {
                behavior: common(None, false),
                ..valid.clone()
            },
            TimeSetBehavior {
                behavior: common(Some(vec!["image".to_string()]), true),
                ..valid.clone()
            },
            TimeSetBehavior {
                behavior: common(Some(vec!["ppt_x".to_string(), "ppt_y".to_string()]), true),
                ..valid.clone()
            },
        ] {
            assert!(write_time_set_behavior(&invalid).is_err());
        }
        for value in ["", "-", ".", "1-", "1e-2", "(unknown+1)"] {
            let invalid = TimeSetBehavior {
                atom: TimeSetBehaviorAtom {
                    to_used: true,
                    value_type: None,
                },
                to: Some(value.to_string()),
                behavior: common(Some(vec!["ppt_x".to_string()]), true),
            };
            assert!(write_time_set_behavior(&invalid).is_err(), "{value}");
        }
        for value in ["123456", "#12345", "#12345g", "#1234567"] {
            let invalid = TimeSetBehavior {
                atom: TimeSetBehaviorAtom {
                    to_used: true,
                    value_type: Some(TimeAnimateValueType::Color),
                },
                to: Some(value.to_string()),
                behavior: common(Some(vec!["fill.color".to_string()]), true),
            };
            assert!(write_time_set_behavior(&invalid).is_err(), "{value}");
        }

        let mut atom = write_time_set_behavior_atom(&TimeSetBehaviorAtom {
            to_used: false,
            value_type: None,
        });
        atom[12..16].copy_from_slice(&2u32.to_le_bytes());
        let (record, _) = PptRecord::parse(&atom, 0).unwrap();
        assert!(parse_time_set_behavior_atom(&record).is_err());
    }

    #[test]
    fn round_trips_and_validates_command_behaviors() {
        assert_eq!(PptRecordType::TimeCommandBehaviorContainer.as_u16(), 0xF132);
        assert_eq!(PptRecordType::TimeCommandBehavior.as_u16(), 0xF13B);
        let common = || TimeBehavior {
            atom: TimeBehaviorAtom {
                additive: None,
                attribute_names_used: false,
            },
            attribute_names: None,
            properties: None,
            target: TimeVisualElement::Sound {
                kind: TimeVisualElementKind::Audio,
                sound_id_ref: 4,
            },
        };
        for (command_type, command) in [
            (Some(TimeCommandBehaviorType::Event), "onstopaudio"),
            (Some(TimeCommandBehaviorType::Call), "playFrom(1.25)"),
            (Some(TimeCommandBehaviorType::OleVerb), "-2"),
            (None, "togglePause"),
        ] {
            let expected = TimeCommandBehavior {
                atom: TimeCommandBehaviorAtom {
                    command_type,
                    command_used: true,
                },
                command: Some(command.to_string()),
                behavior: common(),
            };
            let atom_bytes = write_time_command_behavior_atom(&expected.atom);
            let (atom_record, _) = PptRecord::parse(&atom_bytes, 0).unwrap();
            assert_eq!(
                parse_time_command_behavior_atom(&atom_record).unwrap(),
                expected.atom
            );
            let bytes = write_time_command_behavior(&expected).unwrap();
            let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
            assert_eq!(parse_time_command_behavior(&record).unwrap(), expected);
        }

        for (command_type, command) in [
            (TimeCommandBehaviorType::Event, "stop"),
            (TimeCommandBehaviorType::Call, "playFrom(-1)"),
            (TimeCommandBehaviorType::OleVerb, "verb"),
        ] {
            let invalid = TimeCommandBehavior {
                atom: TimeCommandBehaviorAtom {
                    command_type: Some(command_type),
                    command_used: true,
                },
                command: Some(command.to_string()),
                behavior: common(),
            };
            assert!(write_time_command_behavior(&invalid).is_err());
        }

        let mut atom = write_time_command_behavior_atom(&TimeCommandBehaviorAtom {
            command_type: None,
            command_used: false,
        });
        atom[12..16].copy_from_slice(&0u32.to_le_bytes());
        let (record, _) = PptRecord::parse(&atom, 0).unwrap();
        assert!(parse_time_command_behavior_atom(&record).is_err());
    }

    #[test]
    fn round_trips_iterate_and_sequence_data_atoms() {
        assert_eq!(PptRecordType::TimeIterateData.as_u16(), 0xF140);
        assert_eq!(PptRecordType::TimeSequenceData.as_u16(), 0xF141);
        for iterate_type in [
            TimeIterateType::AllAtOnce,
            TimeIterateType::ByWord,
            TimeIterateType::ByLetter,
        ] {
            for direction in [
                TimeIterateDirection::Backward,
                TimeIterateDirection::Forward,
            ] {
                for interval_type in [
                    TimeIterateIntervalType::Milliseconds,
                    TimeIterateIntervalType::TenthsOfAPercent,
                ] {
                    let expected = TimeIterateData {
                        interval: Some(250),
                        iterate_type: Some(iterate_type),
                        direction: Some(direction),
                        interval_type: Some(interval_type),
                    };
                    let bytes = write_time_iterate_data(&expected);
                    let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
                    assert_eq!(parse_time_iterate_data(&record).unwrap(), expected);
                }
            }
        }
        let expected = TimeSequenceData {
            concurrent: Some(true),
            next_action: Some(TimeSequenceNextAction::SeekToNaturalEnd),
            previous_action: Some(TimeSequencePreviousAction::SkipTimedChildren),
        };
        let bytes = write_time_sequence_data(&expected);
        let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(parse_time_sequence_data(&record).unwrap(), expected);

        let mut bytes = write_time_iterate_data(&TimeIterateData {
            interval: None,
            iterate_type: None,
            direction: None,
            interval_type: None,
        });
        bytes[8] = 1;
        let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
        assert!(parse_time_iterate_data(&record).is_err());

        let mut bytes = write_time_sequence_data(&TimeSequenceData {
            concurrent: Some(false),
            next_action: None,
            previous_action: None,
        });
        bytes[8] = 2;
        let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
        assert!(parse_time_sequence_data(&record).is_err());
    }

    #[test]
    fn round_trips_time_conditions_and_modifiers() {
        assert_eq!(PptRecordType::TimeConditionContainer.as_u16(), 0xF125);
        assert_eq!(PptRecordType::TimeCondition.as_u16(), 0xF128);
        assert_eq!(PptRecordType::TimeModifier.as_u16(), 0xF129);
        let condition_types = [
            TimeConditionType::None,
            TimeConditionType::Begin,
            TimeConditionType::End,
            TimeConditionType::Next,
            TimeConditionType::Previous,
            TimeConditionType::EndSync,
        ];
        for condition_type in condition_types {
            let expected = TimeCondition {
                condition_type,
                atom: TimeConditionAtom {
                    trigger_object: TimeTriggerObject::VisualElement,
                    trigger_event: TimeTriggerEvent::MouseClick,
                    target_id: 0,
                    delay_ms: -1,
                },
                visual_target: Some(TimeVisualElement::Page),
            };
            let bytes = write_time_condition(&expected).unwrap();
            let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
            assert_eq!(consumed, bytes.len());
            assert_eq!(parse_time_condition(&record).unwrap(), expected);
        }

        for trigger_event in [
            TimeTriggerEvent::None,
            TimeTriggerEvent::OnBegin,
            TimeTriggerEvent::TimeNodeStart,
            TimeTriggerEvent::TimeNodeEnd,
            TimeTriggerEvent::MouseClick,
            TimeTriggerEvent::MouseOver,
            TimeTriggerEvent::OnNext,
            TimeTriggerEvent::OnPrevious,
            TimeTriggerEvent::StopAudio,
        ] {
            let atom = TimeConditionAtom {
                trigger_object: TimeTriggerObject::RuntimeNodeReference,
                trigger_event,
                target_id: 2,
                delay_ms: i32::MIN,
            };
            let bytes = write_time_condition_atom(&atom).unwrap();
            let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
            assert_eq!(parse_time_condition_atom(&record).unwrap(), atom);
        }

        let modifiers = [
            TimeModifier::RepeatCount(1),
            TimeModifier::RepeatDuration(2),
            TimeModifier::Speed(3),
            TimeModifier::Accelerate(4),
            TimeModifier::Decelerate(5),
            TimeModifier::AutomaticReverse(u32::MAX),
        ];
        for modifier in modifiers {
            let bytes = write_time_modifier(&modifier);
            let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
            assert_eq!(parse_time_modifier(&record).unwrap(), modifier);
        }
    }

    #[test]
    fn rejects_malformed_time_conditions_and_modifiers() {
        let missing_target = TimeCondition {
            condition_type: TimeConditionType::Begin,
            atom: TimeConditionAtom {
                trigger_object: TimeTriggerObject::VisualElement,
                trigger_event: TimeTriggerEvent::OnBegin,
                target_id: 0,
                delay_ms: 0,
            },
            visual_target: None,
        };
        assert!(write_time_condition(&missing_target).is_err());
        let bad_runtime = TimeConditionAtom {
            trigger_object: TimeTriggerObject::RuntimeNodeReference,
            trigger_event: TimeTriggerEvent::TimeNodeStart,
            target_id: 1,
            delay_ms: 0,
        };
        assert!(write_time_condition_atom(&bad_runtime).is_err());

        let mut bytes = write_time_condition_atom(&TimeConditionAtom {
            trigger_object: TimeTriggerObject::None,
            trigger_event: TimeTriggerEvent::None,
            target_id: 0,
            delay_ms: 0,
        })
        .unwrap();
        bytes[12..16].copy_from_slice(&2u32.to_le_bytes());
        let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
        assert!(parse_time_condition_atom(&record).is_err());

        let mut bytes = write_time_modifier(&TimeModifier::Speed(100));
        bytes[8..12].copy_from_slice(&6u32.to_le_bytes());
        let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
        assert!(parse_time_modifier(&record).is_err());
    }

    #[test]
    fn parses_slide_animation_extensions_and_rejects_duplicates_or_truncation() {
        let node = ExtendedTimeNode {
            atom: TimeNodeAtom {
                node_type: Some(TimeNodeKind::Sequential),
                ..TimeNodeAtom::default()
            },
            properties: None,
            children: Vec::new(),
        };
        let build_list = BuildList::new();
        let mut data = Vec::new();
        data.extend(0u16.to_le_bytes());
        data.extend(0x7777u16.to_le_bytes());
        data.extend(0u32.to_le_bytes());
        data.extend(write_extended_time_node(&node).unwrap());
        let build_bytes = write_build_list(&build_list).unwrap();
        data.extend(&build_bytes);

        let parsed = parse_slide_animation_extension(&data).unwrap();
        assert_eq!(parsed.time_node, Some(node));
        assert_eq!(parsed.build_list, Some(build_list));

        let mut duplicate = data.clone();
        duplicate.extend(build_bytes);
        assert!(parse_slide_animation_extension(&duplicate).is_err());

        let mut truncated = data;
        truncated.pop();
        assert!(parse_slide_animation_extension(&truncated).is_err());
    }

    #[test]
    fn round_trips_exact_powerpoint_2002_build_lists() {
        assert_eq!(PptRecordType::BuildList.as_u16(), 0x2B02);
        assert_eq!(PptRecordType::LevelInfoAtom.as_u16(), 0x2B0A);
        let bytes = write_build_list(&sample_build_list()).unwrap();
        assert_eq!(PptRecordType::TimeNode.as_u16(), 0xF127);
        assert_eq!(bytes.len(), 216);
        let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(record.record_type, PptRecordType::BuildList);
        assert_eq!(record.children.len(), 3);

        let parsed = parse_build_list(&record).unwrap();
        assert_eq!(parsed.builds.len(), 3);
        let BuildListEntry::Paragraph(paragraph) = &parsed.builds[0] else {
            panic!("expected paragraph build");
        };
        assert_eq!(paragraph.atom.build_id, 10);
        assert_eq!(paragraph.atom.shape_id_ref, 100);
        assert_eq!(paragraph.paragraph.build_type, ParagraphBuildType::AsAWhole);
        assert_eq!(paragraph.paragraph.delay_time_ms, 750);
        assert_eq!(paragraph.levels.len(), 1);
        assert_eq!(paragraph.levels[0].level, 0);
        assert_eq!(paragraph.levels[0].time_node.atom, TimeNodeAtom::default());

        let BuildListEntry::Chart(chart) = &parsed.builds[1] else {
            panic!("expected chart build");
        };
        assert_eq!(chart.chart.build_type, ChartBuildType::ByElementInCategory);
        assert!(chart.chart.animate_background);

        let BuildListEntry::Diagram(diagram) = &parsed.builds[2] else {
            panic!("expected diagram build");
        };
        assert_eq!(
            diagram.diagram.build_type,
            DiagramBuildType::CounterClockwiseOut
        );
    }

    #[test]
    fn rejects_malformed_powerpoint_2002_build_lists() {
        let bytes = write_build_list(&sample_build_list()).unwrap();
        let (valid, _) = PptRecord::parse(&bytes, 0).unwrap();

        let mut truncated = bytes.clone();
        let claimed_length = u32::from_le_bytes(truncated[4..8].try_into().unwrap()) + 1;
        truncated[4..8].copy_from_slice(&claimed_length.to_le_bytes());
        let (truncated, _) = PptRecord::parse(&truncated, 0).unwrap();
        assert_eq!(truncated.data_length, claimed_length);
        assert!(parse_build_list(&truncated).is_err());

        let mut malformed = Vec::new();
        let mut wrong_header = valid.clone();
        wrong_header.version = 0;
        malformed.push(wrong_header);

        let mut wrong_bool = valid.clone();
        wrong_bool.children[1].children[1].data[4] = 2;
        malformed.push(wrong_bool);

        let mut wrong_kind = valid.clone();
        wrong_kind.children[2].children[0].data[0..4].copy_from_slice(&1u32.to_le_bytes());
        malformed.push(wrong_kind);

        let mut wrong_level = valid.clone();
        wrong_level.children[0].children[2].data[0..4].copy_from_slice(&10u32.to_le_bytes());
        malformed.push(wrong_level);

        let mut duplicate = valid.clone();
        duplicate.children[2].children[0].data[4..8].copy_from_slice(&11u32.to_le_bytes());
        duplicate.children[2].children[0].data[8..12].copy_from_slice(&101u32.to_le_bytes());
        malformed.push(duplicate);

        let mut wrong_order = valid.clone();
        wrong_order.children[1].children.swap(0, 1);
        malformed.push(wrong_order);

        for record in malformed {
            assert!(parse_build_list(&record).is_err());
        }
    }

    #[test]
    fn test_animation_info_default() {
        let info = AnimationInfo::default();
        assert!(!info.has_animations());
        assert_eq!(info.animation_count(), 0);
    }

    #[test]
    fn test_build_info_default() {
        let build_info = BuildInfo::default();
        assert!(build_info.builds.is_empty());
    }
}
