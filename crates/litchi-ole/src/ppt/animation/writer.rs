//! Animation record writer.
//!
//! Writes PowerPoint binary animation records from structured types.

use super::types::{
    AfterEffect, AnimationEffect, AnimationInfo, BuildAtom, BuildKind, BuildList, BuildListEntry,
    ChartBuild, DiagramBuild, EffectDirection, ExtendedTimeNode, LegacyAnimationAtom,
    LegacyAnimationBuild, LegacyAnimationEffect, LegacyTextBuildSubEffect, ParagraphBuild,
    ParagraphBuildLevel, TimeAnimateColor, TimeAnimateColorBy, TimeBehavior, TimeBehaviorAdditive,
    TimeBehaviorAtom, TimeBehaviorProperty, TimeBehaviorPropertyList, TimeColorBehavior,
    TimeColorBehaviorAtom, TimeColorDirection, TimeColorModel, TimeCommandBehavior,
    TimeCommandBehaviorAtom, TimeCommandBehaviorType, TimeCondition, TimeConditionAtom,
    TimeConditionType, TimeEffectBehavior, TimeEffectBehaviorAtom, TimeEffectNodeType,
    TimeEffectTransition, TimeEffectType, TimeIterateData, TimeIterateDirection,
    TimeIterateIntervalType, TimeIterateType, TimeMasterRelation, TimeModifier, TimeNodeAtom,
    TimeNodeProperty, TimeNodePropertyList, TimePropertyListContext, TimeRotationBehavior,
    TimeRotationBehaviorAtom, TimeRotationDirection, TimeScaleBehavior, TimeScaleBehaviorAtom,
    TimeSequenceData, TimeSequenceNextAction, TimeSequencePreviousAction, TimeTriggerEvent,
    TimeTriggerObject, TimeVisualElement, TimeVisualElementKind, is_valid_runtime_context,
    is_valid_time_filter, is_valid_time_points_types,
};
use crate::consts::PptRecordType;
use crate::ppt::package::{PptError, Result};

/// Write InteractiveInfo container with InteractiveInfoAtom for animations.
/// Per POI MovieShape, this is required alongside AnimationInfo in ClientData.
/// For sound animations, soundRef should match AnimationInfoAtom.soundRef
pub fn write_interactive_info_with_sound(sound_ref: u32) -> Vec<u8> {
    let mut data = Vec::new();

    // InteractiveInfoAtom (16 bytes)
    let mut atom_data: Vec<u8> = Vec::new();
    atom_data.extend(&sound_ref.to_le_bytes()); // soundRef - matches AnimationInfoAtom.soundRef for sounds
    atom_data.extend(&0u32.to_le_bytes()); // exHyperlinkIdRef
    atom_data.extend(&6u8.to_le_bytes()); // action = ACTION_MEDIA per MovieShape
    atom_data.extend(&0u8.to_le_bytes()); // oleVerb
    atom_data.extend(&0u8.to_le_bytes()); // jump
    atom_data.extend(&0u8.to_le_bytes()); // flags
    atom_data.extend(&9u8.to_le_bytes()); // hyperlinkType = LINK_NULL per MovieShape
    atom_data.extend(&0u8.to_le_bytes()); // unknown1
    atom_data.extend(&0u8.to_le_bytes()); // unknown2
    atom_data.extend(&0u8.to_le_bytes()); // unknown3

    let atom_header = create_record_header(
        PptRecordType::InteractiveInfoAtom,
        0x00,
        0,
        atom_data.len() as u32,
    );

    let mut children = Vec::new();
    children.extend(atom_header);
    children.extend(atom_data);

    // InteractiveInfo container wrapping the atom
    let header = create_record_header(
        PptRecordType::InteractiveInfo,
        0x0F,
        0,
        children.len() as u32,
    );
    data.extend(header);
    data.extend(children);

    data
}

/// Write AnimationInfo container record.
/// Returns (AnimationInfo bytes, sound_ref for InteractiveInfo)
pub fn write_animation_info(info: &AnimationInfo) -> Result<(Vec<u8>, u32)> {
    if !info.time_nodes.is_empty() {
        return Err(PptError::InvalidFormat(
            "extended time nodes belong to the slide animation extension, not AnimationInfo"
                .to_string(),
        ));
    }
    let mut data = Vec::new();

    let mut children: Vec<u8> = Vec::new();

    // AnimationInfoAtom MUST be the first child (per POI)
    // Extract first build item to determine animation type and sound
    let atom = if let Some(atom) = &info.legacy_atom {
        atom.clone()
    } else {
        let (fly_method, fly_direction, sound_ref, has_sound) = if let Some(ref build_list) =
            info.build_list
        {
            if let Some(first_build) = build_list.builds.first() {
                let (method, dir) = map_effect_to_ppt97(first_build.effect, first_build.direction);
                let (snd_ref, has_snd) = if let Some(ref sound) = first_build.sound {
                    (sound.sound_ref, true)
                } else {
                    (0, false)
                };
                (method, dir, snd_ref, has_snd)
            } else {
                (0x00, 0, 0, false)
            }
        } else {
            (0x00, 0, 0, false)
        };
        LegacyAnimationAtom {
            has_sound,
            sound_id_ref: sound_ref,
            build_type: if info.has_animations() {
                LegacyAnimationBuild::OneBuild
            } else {
                LegacyAnimationBuild::NoBuild
            },
            effect: LegacyAnimationEffect::parse(fly_method).unwrap_or_default(),
            effect_direction: fly_direction,
            text_build_sub_effect: match info.iteration {
                super::triggers::IterationType::ByWord => LegacyTextBuildSubEffect::ByWord,
                super::triggers::IterationType::ByLetter => LegacyTextBuildSubEffect::ByCharacter,
                _ => LegacyTextBuildSubEffect::AllAtOnce,
            },
            ..LegacyAnimationAtom::default()
        }
    };
    let sound_ref = atom.sound_id_ref;
    children.extend(write_animation_info_atom(&atom)?);

    // NOTE: BuildList is omitted for ClientData embedding per POI AnimationInfo constructor
    // POI AnimationInfo contains ONLY AnimationInfoAtom when embedded in shape ClientData
    // BuildList would be at slide level for multi-shape animations, not per-shape
    // if let Some(ref build_list) = info.build_list {
    //     children.extend(write_build_list(build_list));
    // }

    for raw_record in &info.raw_records {
        if raw_record.record_type == PptRecordType::AnimationInfoAtom {
            return Err(PptError::InvalidFormat(
                "raw AnimationInfo children cannot contain another AnimationInfoAtom".to_string(),
            ));
        }
        children.extend(serialize_raw_record(raw_record));
    }

    let header = create_record_header(PptRecordType::AnimationInfo, 0x0F, 0, children.len() as u32);
    data.extend(header);
    data.extend(children);

    Ok((data, sound_ref))
}

/// Serialize an exact PowerPoint 97 `AnimationInfoAtom`.
pub fn write_animation_info_atom(atom: &LegacyAnimationAtom) -> Result<Vec<u8>> {
    if atom.automatic && atom.delay_time_ms < 0 {
        return Err(PptError::InvalidFormat(
            "automatic AnimationInfoAtom cannot have a negative delay".to_string(),
        ));
    }
    if atom.order_id < -2 {
        return Err(PptError::InvalidFormat(format!(
            "AnimationInfoAtom orderID {} is less than -2",
            atom.order_id
        )));
    }
    if !atom.effect.accepts_direction(atom.effect_direction) {
        return Err(PptError::InvalidFormat(format!(
            "AnimationInfoAtom direction {:#04X} is invalid for {:?}",
            atom.effect_direction, atom.effect
        )));
    }

    let mut data = Vec::with_capacity(28);
    data.extend(atom.dim_color.to_le_bytes());
    let flags = [
        atom.reverse,
        atom.automatic,
        atom.has_sound,
        atom.stop_sound,
        atom.play,
        atom.synchronous,
        atom.hide_while_not_playing,
        atom.animate_background,
    ]
    .into_iter()
    .enumerate()
    .fold(0u16, |flags, (index, value)| {
        flags | (u16::from(value) << (index * 2))
    });
    data.extend(flags.to_le_bytes());
    data.extend(0u16.to_le_bytes());
    data.extend(atom.sound_id_ref.to_le_bytes());
    data.extend(atom.delay_time_ms.to_le_bytes());
    data.extend(atom.order_id.to_le_bytes());
    data.extend(atom.slide_count.to_le_bytes());
    data.push(atom.build_type.as_u8());
    data.push(atom.effect.as_u8());
    data.push(atom.effect_direction);
    data.push(match atom.after_effect {
        AfterEffect::None => 0,
        AfterEffect::DimToColor => 1,
        AfterEffect::HideOnNextClick => 2,
        AfterEffect::Hide => 3,
    });
    data.push(atom.text_build_sub_effect.as_u8());
    data.push(atom.ole_verb);
    data.extend([0, 0]);

    let mut result = create_record_header(PptRecordType::AnimationInfoAtom, 0x01, 0, 28);
    result.extend(data);
    Ok(result)
}

/// Serialize an exact PowerPoint 2002 extended time-node envelope.
pub fn write_extended_time_node(node: &ExtendedTimeNode) -> Result<Vec<u8>> {
    let mut children = write_time_node_atom(&node.atom);
    if let Some(properties) = &node.properties {
        children.extend(write_time_node_property_list(
            properties,
            TimePropertyListContext::TimeNode,
        )?);
    }
    for child in &node.children {
        if matches!(
            child.record_type,
            PptRecordType::TimeNode | PptRecordType::TimePropertyList
        ) {
            return Err(PptError::InvalidFormat(
                "extended time-node raw children cannot contain atom/property-list records"
                    .to_string(),
            ));
        }
        if child.data_length as usize != child.data.len() {
            return Err(PptError::InvalidFormat(format!(
                "extended time-node child {:?} has a truncated payload",
                child.record_type
            )));
        }
        children.extend(serialize_raw_record(child));
    }
    wrap_record(PptRecordType::ExtTimeNode, 0x0F, 1, children)
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
    let mut result = create_record_header(PptRecordType::TimeNode, 0, 0, 32);
    result.extend(data);
    result
}

/// Serialize a typed `TimePropertyList4TimeNodeContainer`.
pub fn write_time_node_property_list(
    list: &TimeNodePropertyList,
    context: TimePropertyListContext,
) -> Result<Vec<u8>> {
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
            return Err(PptError::InvalidFormat(format!(
                "duplicate time property {id:#X}"
            )));
        }
        validate_time_property_context(id, context)?;
        if matches!(property, TimeNodeProperty::EventFilter(_)) && !has_interactive_sequence {
            return Err(PptError::InvalidFormat(
                "event filter requires an interactive sequence".to_string(),
            ));
        }
        let length = u32::try_from(data.len()).map_err(|_| {
            PptError::InvalidFormat("time property exceeds 4 GiB record limit".to_string())
        })?;
        children.extend(create_record_header(
            PptRecordType::TimeVariant,
            0,
            id,
            length,
        ));
        children.extend(data);
    }
    wrap_record(PptRecordType::TimePropertyList, 0x0F, 0, children)
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
                return Err(PptError::InvalidFormat("invalid time filter".to_string()));
            }
            (0x10, string(value))
        },
        TimeNodeProperty::EventFilter(value) => {
            if value != "cancelBubble" {
                return Err(PptError::InvalidFormat(
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
                return Err(PptError::InvalidFormat(
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
        return Err(PptError::InvalidFormat(format!(
            "time property {id:#X} is invalid in {context:?} context"
        )));
    }
    Ok(())
}

/// Serialize the common behavior information shared by extended animation behaviors.
pub fn write_time_behavior(behavior: &TimeBehavior) -> Result<Vec<u8>> {
    let mut children = write_time_behavior_atom(&behavior.atom);
    if let Some(attribute_names) = &behavior.attribute_names {
        children.extend(write_time_string_list(attribute_names)?);
    }
    if let Some(properties) = &behavior.properties {
        children.extend(write_time_behavior_property_list(properties)?);
    }
    children.extend(write_time_visual_element(&behavior.target)?);
    wrap_record(PptRecordType::TimeBehaviorContainer, 0x0F, 0, children)
}

/// Serialize an exact 16-byte `TimeBehaviorAtom` payload.
pub fn write_time_behavior_atom(atom: &TimeBehaviorAtom) -> Vec<u8> {
    let mut data = Vec::with_capacity(16);
    let flags = u32::from(atom.additive.is_some()) | (u32::from(atom.attribute_names_used) << 2);
    data.extend(flags.to_le_bytes());
    data.extend(
        atom.additive
            .map_or(0u32, |value| match value {
                TimeBehaviorAdditive::Override => 0,
                TimeBehaviorAdditive::Add => 1,
            })
            .to_le_bytes(),
    );
    data.extend(0u32.to_le_bytes());
    data.extend(0u32.to_le_bytes());
    let mut result = create_record_header(PptRecordType::TimeBehavior, 0, 0, 16);
    result.extend(data);
    result
}

/// Serialize an exact color behavior container.
pub fn write_time_color_behavior(behavior: &TimeColorBehavior) -> Result<Vec<u8>> {
    validate_color_behavior(&behavior.atom, &behavior.behavior)?;
    let mut children = write_time_color_behavior_atom(&behavior.atom)?;
    children.extend(write_time_behavior(&behavior.behavior)?);
    wrap_record(PptRecordType::TimeColorBehaviorContainer, 0x0F, 0, children)
}

/// Serialize an exact `TimeColorBehaviorAtom`.
pub fn write_time_color_behavior_atom(atom: &TimeColorBehaviorAtom) -> Result<Vec<u8>> {
    if atom.from.is_some() && atom.by.is_none() && atom.to.is_none() {
        return Err(PptError::InvalidFormat(
            "color from value requires a by or to value".to_string(),
        ));
    }
    let flags = u32::from(atom.by.is_some())
        | (u32::from(atom.from.is_some()) << 1)
        | (u32::from(atom.to.is_some()) << 2)
        | (u32::from(atom.color_space_used) << 3)
        | (u32::from(atom.direction_used) << 4);
    let mut data = Vec::with_capacity(52);
    data.extend(flags.to_le_bytes());
    data.extend(match &atom.by {
        Some(color) => encode_animate_color_by(color)?,
        None => [0; 16],
    });
    data.extend(match &atom.from {
        Some(color) => encode_animate_color(color)?,
        None => [0; 16],
    });
    data.extend(match &atom.to {
        Some(color) => encode_animate_color(color)?,
        None => [0; 16],
    });
    let mut result = create_record_header(PptRecordType::TimeColorBehavior, 0, 0, 52);
    result.extend(data);
    Ok(result)
}

fn encode_animate_color_by(color: &TimeAnimateColorBy) -> Result<[u8; 16]> {
    let (model, values) = match color {
        TimeAnimateColorBy::Rgb { red, green, blue } => (0u32, [*red, *green, *blue]),
        TimeAnimateColorBy::Hsl {
            hue,
            saturation,
            luminance,
        } => (1, [*hue, *saturation, *luminance]),
        TimeAnimateColorBy::Scheme(index) => return encode_scheme_color(*index),
    };
    if values.iter().any(|value| !(-255..=255).contains(value)) {
        return Err(PptError::InvalidFormat(
            "color offset component is out of range".to_string(),
        ));
    }
    let mut data = [0; 16];
    data[0..4].copy_from_slice(&model.to_le_bytes());
    for (index, value) in values.into_iter().enumerate() {
        data[4 + index * 4..8 + index * 4].copy_from_slice(&value.to_le_bytes());
    }
    Ok(data)
}

fn encode_animate_color(color: &TimeAnimateColor) -> Result<[u8; 16]> {
    match color {
        TimeAnimateColor::Scheme(index) => encode_scheme_color(*index),
        TimeAnimateColor::Rgb { red, green, blue } => {
            if *red > 255 || *green > 255 || *blue > 255 {
                return Err(PptError::InvalidFormat(
                    "RGB color component is out of range".to_string(),
                ));
            }
            let mut data = [0; 16];
            for (index, value) in [*red, *green, *blue].into_iter().enumerate() {
                data[4 + index * 4..8 + index * 4].copy_from_slice(&value.to_le_bytes());
            }
            Ok(data)
        },
    }
}

fn encode_scheme_color(index: u32) -> Result<[u8; 16]> {
    if index > 7 {
        return Err(PptError::InvalidFormat(
            "scheme color index is out of range".to_string(),
        ));
    }
    let mut data = [0; 16];
    data[0..4].copy_from_slice(&2u32.to_le_bytes());
    data[4..8].copy_from_slice(&index.to_le_bytes());
    Ok(data)
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

/// Serialize an exact image-effect behavior container.
pub fn write_time_effect_behavior(effect: &TimeEffectBehavior) -> Result<Vec<u8>> {
    validate_effect_behavior(effect)?;
    let mut children = write_time_effect_behavior_atom(&effect.atom);
    if let Some(filter) = effect.filter {
        append_time_variant(
            &mut children,
            1,
            encode_time_variant_string(filter.as_str()),
        )?;
    }
    if let Some(progress) = effect.progress {
        let mut data = vec![2];
        data.extend(progress.to_le_bytes());
        append_time_variant(&mut children, 2, data)?;
    }
    if let Some(runtime_context) = &effect.runtime_context {
        append_time_variant(
            &mut children,
            3,
            encode_time_variant_string(runtime_context),
        )?;
    }
    children.extend(write_time_behavior(&effect.behavior)?);
    wrap_record(
        PptRecordType::TimeEffectBehaviorContainer,
        0x0F,
        0,
        children,
    )
}

/// Serialize an exact `TimeEffectBehaviorAtom`.
pub fn write_time_effect_behavior_atom(atom: &TimeEffectBehaviorAtom) -> Vec<u8> {
    let flags = u32::from(atom.transition.is_some())
        | (u32::from(atom.filter_used) << 1)
        | (u32::from(atom.progress_used) << 2)
        | (u32::from(atom.runtime_context_used) << 3);
    let transition = atom.transition.map_or(0u32, |value| match value {
        TimeEffectTransition::In => 0,
        TimeEffectTransition::Out => 1,
    });
    let mut result = create_record_header(PptRecordType::TimeEffectBehavior, 0, 0, 8);
    result.extend(flags.to_le_bytes());
    result.extend(transition.to_le_bytes());
    result
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

fn append_time_variant(children: &mut Vec<u8>, instance: u16, data: Vec<u8>) -> Result<()> {
    let length = u32::try_from(data.len())
        .map_err(|_| PptError::InvalidFormat("time variant exceeds 4 GiB".to_string()))?;
    children.extend(create_record_header(
        PptRecordType::TimeVariant,
        0,
        instance,
        length,
    ));
    children.extend(data);
    Ok(())
}

/// Serialize a typed `TimePropertyList4TimeBehavior` record.
pub fn write_time_behavior_property_list(list: &TimeBehaviorPropertyList) -> Result<Vec<u8>> {
    let mut seen = std::collections::HashSet::with_capacity(list.properties.len());
    let mut children = Vec::new();
    for property in &list.properties {
        let (id, data) = encode_time_behavior_property(property)?;
        if !seen.insert(id) {
            return Err(PptError::InvalidFormat(format!(
                "duplicate time behavior property {id:#X}"
            )));
        }
        let length = u32::try_from(data.len()).map_err(|_| {
            PptError::InvalidFormat("time behavior property exceeds 4 GiB".to_string())
        })?;
        children.extend(create_record_header(
            PptRecordType::TimeVariant,
            0,
            id,
            length,
        ));
        children.extend(data);
    }
    wrap_record(PptRecordType::TimePropertyList, 0x0F, 0, children)
}

fn encode_time_behavior_property(property: &TimeBehaviorProperty) -> Result<(u16, Vec<u8>)> {
    let integer = |value: i32| {
        let mut data = vec![1];
        data.extend(value.to_le_bytes());
        data
    };
    let float = |value: f32| {
        let mut data = vec![2];
        data.extend(value.to_le_bytes());
        data
    };
    let string = |value: &str| encode_time_variant_string(value);
    Ok(match property {
        TimeBehaviorProperty::UnknownPropertyList(value) => (0x01, string(value)),
        TimeBehaviorProperty::RuntimeContext(value) => {
            if !is_valid_runtime_context(value) {
                return Err(PptError::InvalidFormat(
                    "invalid time runtime context".to_string(),
                ));
            }
            (0x02, string(value))
        },
        TimeBehaviorProperty::MotionPathEditRelative(value) => (0x03, vec![0, u8::from(*value)]),
        TimeBehaviorProperty::ColorModel(value) => (
            0x04,
            integer(match value {
                TimeColorModel::Rgb => 0,
                TimeColorModel::Hsl => 1,
                TimeColorModel::Scheme => 2,
            }),
        ),
        TimeBehaviorProperty::ColorDirection(value) => (
            0x05,
            integer(match value {
                TimeColorDirection::Clockwise => 0,
                TimeColorDirection::CounterClockwise => 1,
            }),
        ),
        TimeBehaviorProperty::Override => (0x06, integer(1)),
        TimeBehaviorProperty::PathEditRotationAngle(value) => (0x07, float(*value)),
        TimeBehaviorProperty::PathEditRotationX(value) => (0x08, float(*value)),
        TimeBehaviorProperty::PathEditRotationY(value) => (0x09, float(*value)),
        TimeBehaviorProperty::PointsTypes(value) => {
            if !is_valid_time_points_types(value) {
                return Err(PptError::InvalidFormat(
                    "invalid time path point types".to_string(),
                ));
            }
            (0x0A, string(value))
        },
    })
}

fn write_time_string_list(names: &[String]) -> Result<Vec<u8>> {
    let mut children = Vec::new();
    for name in names {
        let data = encode_time_variant_string(name);
        let length = u32::try_from(data.len()).map_err(|_| {
            PptError::InvalidFormat("time attribute name exceeds 4 GiB".to_string())
        })?;
        children.extend(create_record_header(
            PptRecordType::TimeVariant,
            0,
            0,
            length,
        ));
        children.extend(data);
    }
    wrap_record(PptRecordType::TimeVariantList, 0x0F, 1, children)
}

fn encode_time_variant_string(value: &str) -> Vec<u8> {
    let mut data = Vec::with_capacity(1 + value.len().saturating_mul(2));
    data.push(3);
    data.extend(value.encode_utf16().flat_map(u16::to_le_bytes));
    data
}

/// Serialize a `ClientVisualElementContainer` animation target.
pub fn write_time_visual_element(target: &TimeVisualElement) -> Result<Vec<u8>> {
    let atom = match target {
        TimeVisualElement::Page => {
            let mut atom = create_record_header(PptRecordType::VisualPageAtom, 0, 0, 4);
            atom.extend(TimeVisualElementKind::Page.as_u32().to_le_bytes());
            atom
        },
        TimeVisualElement::Sound { kind, sound_id_ref } => {
            if *kind == TimeVisualElementKind::Page {
                return Err(PptError::InvalidFormat(
                    "sound target cannot use the page element type".to_string(),
                ));
            }
            write_visual_shape_atom(*kind, 2, *sound_id_ref, u32::MAX, u32::MAX)
        },
        TimeVisualElement::Shape {
            kind,
            shape_id_ref,
            data1,
            data2,
        } => {
            if matches!(
                kind,
                TimeVisualElementKind::Page | TimeVisualElementKind::ChartElement
            ) {
                return Err(PptError::InvalidFormat(
                    "general shape target has an invalid element type".to_string(),
                ));
            }
            write_visual_shape_atom(
                *kind,
                1,
                *shape_id_ref,
                u32::from_le_bytes(data1.to_le_bytes()),
                u32::from_le_bytes(data2.to_le_bytes()),
            )
        },
        TimeVisualElement::Chart {
            shape_id_ref,
            build_type,
            element_index,
        } => {
            if *element_index < -1 {
                return Err(PptError::InvalidFormat(
                    "chart target element index must be at least -1".to_string(),
                ));
            }
            write_visual_shape_atom(
                TimeVisualElementKind::ChartElement,
                1,
                *shape_id_ref,
                build_type.as_u32(),
                u32::from_le_bytes(element_index.to_le_bytes()),
            )
        },
    };
    wrap_record(PptRecordType::TimeClientVisualElement, 0x0F, 0, atom)
}

fn write_visual_shape_atom(
    kind: TimeVisualElementKind,
    reference_type: u32,
    reference_id: u32,
    data1: u32,
    data2: u32,
) -> Vec<u8> {
    let mut atom = create_record_header(PptRecordType::VisualShapeAtom, 0, 0, 20);
    atom.extend(kind.as_u32().to_le_bytes());
    atom.extend(reference_type.to_le_bytes());
    atom.extend(reference_id.to_le_bytes());
    atom.extend(data1.to_le_bytes());
    atom.extend(data2.to_le_bytes());
    atom
}

/// Serialize an exact rotation behavior container.
pub fn write_time_rotation_behavior(behavior: &TimeRotationBehavior) -> Result<Vec<u8>> {
    validate_rotation_behavior(&behavior.behavior)?;
    let mut children = write_time_rotation_behavior_atom(&behavior.atom)?;
    children.extend(write_time_behavior(&behavior.behavior)?);
    wrap_record(
        PptRecordType::TimeRotationBehaviorContainer,
        0x0F,
        0,
        children,
    )
}

/// Serialize an exact `TimeRotationBehaviorAtom`.
pub fn write_time_rotation_behavior_atom(atom: &TimeRotationBehaviorAtom) -> Result<Vec<u8>> {
    if atom.from_degrees.is_some() && atom.by_degrees.is_none() && atom.to_degrees.is_none() {
        return Err(PptError::InvalidFormat(
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
    let mut result = create_record_header(PptRecordType::TimeRotationBehavior, 0, 0, 20);
    result.extend(data);
    Ok(result)
}

/// Serialize an exact scale behavior container.
pub fn write_time_scale_behavior(behavior: &TimeScaleBehavior) -> Result<Vec<u8>> {
    validate_basic_behavior_properties(&behavior.behavior)?;
    let mut children = write_time_scale_behavior_atom(&behavior.atom)?;
    children.extend(write_time_behavior(&behavior.behavior)?);
    wrap_record(PptRecordType::TimeScaleBehaviorContainer, 0x0F, 0, children)
}

/// Serialize an exact `TimeScaleBehaviorAtom`.
pub fn write_time_scale_behavior_atom(atom: &TimeScaleBehaviorAtom) -> Result<Vec<u8>> {
    if atom.from_percent.is_some() && atom.by_percent.is_none() && atom.to_percent.is_none() {
        return Err(PptError::InvalidFormat(
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
    let mut result = create_record_header(PptRecordType::TimeScaleBehavior, 0, 0, 32);
    result.extend(data);
    Ok(result)
}

fn validate_rotation_behavior(behavior: &TimeBehavior) -> Result<()> {
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
    validate_basic_behavior_properties(behavior)
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
            .map_err(|_| PptError::InvalidFormat("time command exceeds 4 GiB".to_string()))?;
        children.extend(create_record_header(
            PptRecordType::TimeVariant,
            0,
            1,
            length,
        ));
        children.extend(data);
    }
    children.extend(write_time_behavior(&behavior.behavior)?);
    wrap_record(
        PptRecordType::TimeCommandBehaviorContainer,
        0x0F,
        0,
        children,
    )
}

/// Serialize an exact `TimeCommandBehaviorAtom`.
pub fn write_time_command_behavior_atom(atom: &TimeCommandBehaviorAtom) -> Vec<u8> {
    let flags = u32::from(atom.command_type.is_some()) | (u32::from(atom.command_used) << 1);
    let command_type = atom.command_type.map_or(1u32, |value| match value {
        TimeCommandBehaviorType::Event => 0,
        TimeCommandBehaviorType::Call => 1,
        TimeCommandBehaviorType::OleVerb => 2,
    });
    let mut result = create_record_header(PptRecordType::TimeCommandBehavior, 0, 0, 8);
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
        return Err(PptError::InvalidFormat(
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
    let mut result = create_record_header(PptRecordType::TimeIterateData, 0, 0, 20);
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
    let mut result = create_record_header(PptRecordType::TimeSequenceData, 0, 0, 20);
    result.extend(payload);
    result
}

/// Serialize an exact `TimeConditionContainer`.
pub fn write_time_condition(condition: &TimeCondition) -> Result<Vec<u8>> {
    let expects_visual = condition.atom.trigger_object == TimeTriggerObject::VisualElement;
    if expects_visual != condition.visual_target.is_some() {
        return Err(PptError::InvalidFormat(
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
    wrap_record(
        PptRecordType::TimeConditionContainer,
        0x0F,
        instance,
        children,
    )
}

/// Serialize an exact `TimeConditionAtom`.
pub fn write_time_condition_atom(atom: &TimeConditionAtom) -> Result<Vec<u8>> {
    if atom.trigger_object == TimeTriggerObject::RuntimeNodeReference && atom.target_id != 2 {
        return Err(PptError::InvalidFormat(
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
    let mut result = create_record_header(PptRecordType::TimeCondition, 0, 0, 16);
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
    let mut result = create_record_header(PptRecordType::TimeModifier, 0, 0, 8);
    result.extend(kind.to_le_bytes());
    result.extend(value.to_le_bytes());
    result
}

/// Map a high-level animation effect to PPT97 fly method and direction codes.
/// Based on LibreOffice ppt97animations.cxx mapping.
fn map_effect_to_ppt97(effect: AnimationEffect, direction: EffectDirection) -> (u8, u8) {
    use AnimationEffect::*;
    use EffectDirection::*;

    match effect {
        // Entrance effects
        Appear => (0x00, 0),
        FadeIn => (0x06, 0),
        FlyIn => match direction {
            FromLeft => (0x0c, 0x00),
            FromTop => (0x0c, 0x01),
            FromRight => (0x0c, 0x02),
            FromBottom => (0x0c, 0x03),
            FromTopLeft => (0x0c, 0x04),
            FromTopRight => (0x0c, 0x05),
            FromBottomLeft => (0x0c, 0x06),
            FromBottomRight => (0x0c, 0x07),
            _ => (0x0c, 0x00),
        },
        Wipe => match direction {
            FromRight => (0x0a, 0x00),
            FromBottom => (0x0a, 0x01),
            FromLeft => (0x0a, 0x02),
            FromTop => (0x0a, 0x03),
            _ => (0x0a, 0x00),
        },
        Split => (0x0d, 0),
        Dissolve => (0x05, 0),
        Box => match direction {
            Out => (0x0b, 0x00),
            In => (0x0b, 0x01),
            _ => (0x0b, 0x00),
        },
        Checkerboard => match direction {
            Horizontal => (0x03, 0x00),
            Vertical => (0x03, 0x01),
            _ => (0x03, 0x00),
        },
        Blinds => match direction {
            Horizontal => (0x02, 0x00),
            Vertical => (0x02, 0x01),
            _ => (0x02, 0x00),
        },
        RandomBars => match direction {
            Horizontal => (0x08, 0x00),
            Vertical => (0x08, 0x01),
            _ => (0x08, 0x00),
        },
        GrowAndTurn => (0x00, 0),
        // Zoom sub-effects per ppt97animations.cxx:
        // 0x10=zoom-in, 0x11=zoom-in-slightly, 0x12=zoom-out,
        // 0x13=zoom-out-slightly, 0x14=from-screen-center, 0x15=out-from-screen-center
        Zoom => match direction {
            In => (0x0c, 0x10),
            Out => (0x0c, 0x12),
            _ => (0x0c, 0x10),
        },
        Expand => (0x0c, 0x10),   // zoom-in
        Compress => (0x0c, 0x12), // zoom-out
        // Stretch sub-effects: 0x16=across, 0x17=from-left, 0x18=from-top,
        // 0x19=from-right, 0x1a=from-bottom
        Stretch => match direction {
            FromLeft => (0x0c, 0x17),
            FromTop => (0x0c, 0x18),
            FromRight => (0x0c, 0x19),
            FromBottom => (0x0c, 0x1a),
            _ => (0x0c, 0x16),
        },
        // Swivel: 0x1b=vertical
        Swivel => (0x0c, 0x1b),
        // SpiralIn: 0x1c
        SpiralIn => (0x0c, 0x1c),
        Bounce => (0x00, 0),
        // PeekIn sub-effects: 0x08=from-left, 0x09=from-bottom, 0x0a=from-right, 0x0b=from-top
        PeekIn => match direction {
            FromLeft => (0x0c, 0x08),
            FromBottom => (0x0c, 0x09),
            FromRight => (0x0c, 0x0a),
            FromTop => (0x0c, 0x0b),
            _ => (0x0c, 0x08),
        },
        // CrawlIn = slow fly: 0x0c=from-left, 0x0d=from-top, 0x0e=from-right, 0x0f=from-bottom
        CrawlIn => match direction {
            FromLeft => (0x0c, 0x0c),
            FromTop => (0x0c, 0x0d),
            FromRight => (0x0c, 0x0e),
            FromBottom => (0x0c, 0x0f),
            _ => (0x0c, 0x0c),
        },
        FloatIn | Ascend => (0x0c, 0x03), // fly from bottom
        Descend => (0x0c, 0x01),          // fly from top
        RiseUp => (0x0c, 0x03),           // fly from bottom
        Random => (0x01, 0),              // random
        Wheel => (0x1a, 1),
        Plus => (0x12, 0),
        Diamond => (0x11, 0),
        Wedge => (0x13, 0),
        Strips => (0x09, 4),

        // Emphasis effects (map to appear as PPT97 doesn't have these)
        Pulse | Spin | Teeter | Wave | Lighten | Darken => (0x00, 0),
        ChangeFillColor | ChangeLineColor | ChangeFontColor | ChangeFontSize => (0x00, 0),
        GrowShrink | BoldFlash | Underline | ColorPulse => (0x00, 0),
        ComplementaryColor | ComplementaryColor2 | ContrastingColor => (0x00, 0),
        Transparency | ObjectColor | VerticalHighlight | Flicker => (0x00, 0),

        // Exit effects (reverse of entrance)
        FadeOut | Disappear => (0x00, 0),
        FlyOut | WipeOut | BoxOut | CheckerboardOut => (0x00, 0),
        BlindsOut | RandomBarsOut | StripsOut | SplitOut => (0x00, 0),
        PeekOut | PlusOut | DiamondOut | CrawlOut => (0x00, 0),
        DescendOut | Collapse | SinkDown | SpiralOut => (0x00, 0),

        // Motion paths (not supported in PPT97)
        MotionPath | MotionPathLines | MotionPathCurves | MotionPathShapes => (0x00, 0),
        MotionPathLeft | MotionPathRight | MotionPathUp | MotionPathDown => (0x00, 0),
        MotionPathDiagonalUpRight | MotionPathDiagonalDownRight => (0x00, 0),
        MotionPathArcDown | MotionPathArcUp | MotionPathCircle => (0x00, 0),
        MotionPathDiamond | MotionPathHeart | MotionPathHexagon => (0x00, 0),
        MotionPathOctagon | MotionPathPentagon | MotionPathSquare => (0x00, 0),
        MotionPathStar4 | MotionPathStar5 | MotionPathStar6 | MotionPathStar8 => (0x00, 0),
        MotionPathTriangle | MotionPathLoopDeLoop | MotionPathCurvedX => (0x00, 0),
        MotionPathSCurve1 | MotionPathSCurve2 | MotionPathSineWave => (0x00, 0),
        MotionPathSpiralLeft | MotionPathSpiralRight | MotionPathSpring => (0x00, 0),
        MotionPathZigzag => (0x00, 0),

        Custom => (0x00, 0),
    }
}

/// Write BuildList container record.
pub fn write_build_list(build_info: &BuildList) -> Result<Vec<u8>> {
    let mut identities = std::collections::HashSet::with_capacity(build_info.builds.len());
    let mut children = Vec::new();
    for build in &build_info.builds {
        let atom = match build {
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
        children.extend(match build {
            BuildListEntry::Paragraph(build) => write_paragraph_build(build)?,
            BuildListEntry::Chart(build) => write_chart_build(build)?,
            BuildListEntry::Diagram(build) => write_diagram_build(build)?,
        });
    }
    wrap_record(PptRecordType::BuildList, 0x0F, 0, children)
}

fn write_build_atom(atom: &BuildAtom, kind: BuildKind) -> Vec<u8> {
    let mut data = Vec::with_capacity(16);
    data.extend(kind.as_u32().to_le_bytes());
    data.extend(atom.build_id.to_le_bytes());
    data.extend(atom.shape_id_ref.to_le_bytes());
    data.push(u8::from(atom.expanded));
    data.push(u8::from(atom.ui_expanded));
    data.extend([0, 0]);
    let mut result = create_record_header(PptRecordType::BuildAtom, 0, 0, 16);
    result.extend(data);
    result
}

fn write_paragraph_build(build: &ParagraphBuild) -> Result<Vec<u8>> {
    validate_paragraph_levels(&build.paragraph.build_type, &build.levels)?;
    let mut children = write_build_atom(&build.atom, BuildKind::Paragraph);
    let mut atom = Vec::with_capacity(16);
    atom.extend(build.paragraph.build_type.as_u32().to_le_bytes());
    atom.extend(build.paragraph.build_level.to_le_bytes());
    atom.push(u8::from(build.paragraph.animate_background));
    atom.push(u8::from(build.paragraph.reverse));
    atom.push(u8::from(build.paragraph.user_set_animate_background));
    atom.push(u8::from(build.paragraph.automatic));
    atom.extend(build.paragraph.delay_time_ms.to_le_bytes());
    children.extend(create_record_header(PptRecordType::ParaBuildAtom, 1, 0, 16));
    children.extend(atom);
    for level in &build.levels {
        children.extend(create_record_header(PptRecordType::LevelInfoAtom, 0, 0, 4));
        children.extend(level.level.to_le_bytes());
        children.extend(write_extended_time_node(&level.time_node)?);
    }
    wrap_record(PptRecordType::ParaBuild, 0x0F, 0, children)
}

fn write_chart_build(build: &ChartBuild) -> Result<Vec<u8>> {
    let mut children = write_build_atom(&build.atom, BuildKind::Chart);
    let mut atom = Vec::with_capacity(8);
    atom.extend(build.chart.build_type.as_u32().to_le_bytes());
    atom.push(u8::from(build.chart.animate_background));
    atom.extend([0, 0, 0]);
    children.extend(create_record_header(PptRecordType::ChartBuildAtom, 0, 0, 8));
    children.extend(atom);
    wrap_record(PptRecordType::ChartBuild, 0x0F, 0, children)
}

fn write_diagram_build(build: &DiagramBuild) -> Result<Vec<u8>> {
    let mut children = write_build_atom(&build.atom, BuildKind::Diagram);
    children.extend(create_record_header(
        PptRecordType::DiagramBuildAtom,
        0,
        0,
        4,
    ));
    children.extend(build.diagram.build_type.as_u32().to_le_bytes());
    wrap_record(PptRecordType::DiagramBuild, 0x0F, 0, children)
}

fn validate_paragraph_levels(
    build_type: &super::types::ParagraphBuildType,
    levels: &[ParagraphBuildLevel],
) -> Result<()> {
    if levels.is_empty() {
        return Err(PptError::InvalidFormat(
            "ParaBuild requires at least one level".to_string(),
        ));
    }
    if *build_type == super::types::ParagraphBuildType::AsAWhole && levels.len() != 1 {
        return Err(PptError::InvalidFormat(
            "AsAWhole ParaBuild requires exactly one level".to_string(),
        ));
    }
    for (index, level) in levels.iter().enumerate() {
        if level.level > 9 {
            return Err(PptError::InvalidFormat(format!(
                "paragraph build level {} exceeds 9",
                level.level
            )));
        }
        if index > 0 && levels[index - 1].level >= level.level {
            return Err(PptError::InvalidFormat(
                "ParaBuild levels must be strictly increasing".to_string(),
            ));
        }
    }
    Ok(())
}

fn wrap_record(
    record_type: PptRecordType,
    version: u16,
    instance: u16,
    data: Vec<u8>,
) -> Result<Vec<u8>> {
    let length = u32::try_from(data.len()).map_err(|_| {
        PptError::InvalidFormat(format!("{record_type:?} data exceeds 4 GiB record limit"))
    })?;
    let mut result = create_record_header(record_type, version, instance, length);
    result.extend(data);
    Ok(result)
}

/// Create a PPT record header.
fn create_record_header(
    record_type: PptRecordType,
    version: u16,
    instance: u16,
    data_length: u32,
) -> Vec<u8> {
    create_record_header_raw(record_type.as_u16(), version, instance, data_length)
}

fn create_record_header_raw(
    record_type: u16,
    version: u16,
    instance: u16,
    data_length: u32,
) -> Vec<u8> {
    let mut header = Vec::with_capacity(8);

    let version_instance = version | (instance << 4);
    header.extend(&version_instance.to_le_bytes());

    header.extend(&record_type.to_le_bytes());

    header.extend(&data_length.to_le_bytes());

    header
}

/// Serialize raw record (for preserving unknown/complex records).
fn serialize_raw_record(record: &crate::ppt::records::PptRecord) -> Vec<u8> {
    let mut data = Vec::new();

    let header = create_record_header_raw(
        record.record_type_raw,
        record.version,
        record.instance,
        record.data.len() as u32,
    );
    data.extend(header);
    data.extend(&record.data);

    data
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ppt::animation::{ChartBuildAtom, ChartBuildType};

    #[test]
    fn test_write_build_list_empty() {
        let build_info = BuildList::new();
        let data = write_build_list(&build_info).unwrap();

        assert_eq!(data.len(), 8);
    }

    #[test]
    fn rejects_invalid_paragraph_builds() {
        let time_node = ExtendedTimeNode {
            atom: TimeNodeAtom::default(),
            properties: None,
            children: Vec::new(),
        };
        let level = ParagraphBuildLevel {
            level: 10,
            time_node,
        };
        assert!(
            validate_paragraph_levels(
                &super::super::types::ParagraphBuildType::AllAtOnce,
                &[level]
            )
            .is_err()
        );
        assert!(
            validate_paragraph_levels(&super::super::types::ParagraphBuildType::AsAWhole, &[])
                .is_err()
        );
    }

    #[test]
    fn rejects_duplicate_build_id_shape_pairs() {
        let entry = || {
            BuildListEntry::Chart(ChartBuild {
                atom: BuildAtom {
                    build_id: 5,
                    shape_id_ref: 9,
                    expanded: false,
                    ui_expanded: false,
                },
                chart: ChartBuildAtom {
                    build_type: ChartBuildType::AsOneObject,
                    animate_background: false,
                },
            })
        };
        let list = BuildList {
            builds: vec![entry(), entry()],
        };
        assert!(write_build_list(&list).is_err());
    }

    #[test]
    fn rejects_invalid_legacy_animation_atom_combinations() {
        let mut atom = LegacyAnimationAtom {
            effect: LegacyAnimationEffect::Wheel,
            effect_direction: 7,
            ..LegacyAnimationAtom::default()
        };
        assert!(write_animation_info_atom(&atom).is_err());

        atom.effect = LegacyAnimationEffect::Cut;
        atom.effect_direction = 0;
        atom.automatic = true;
        atom.delay_time_ms = -1;
        assert!(write_animation_info_atom(&atom).is_err());
    }
}
