//! Animation record parser.
//!
//! Parses PowerPoint binary animation records into structured types.

use super::triggers::IterationType;
use super::types::{
    AfterEffect, AnimationEffect, AnimationInfo, AnimationTrigger, BuildInfo, BuildLevel,
    BuildType, EffectDirection, EffectSpeed, FillMode, RestartMode, TimeNodeContainer,
    TimeNodeType,
};
use crate::ole::consts::PptRecordType;
use crate::ole::ppt::package::{PptError, Result};
use crate::ole::ppt::records::PptRecord;
use zerocopy::{FromBytes, byteorder::LittleEndian, byteorder::U32};

/// Parse animation info from AnimationInfo container record.
pub fn parse_animation_info(record: &PptRecord) -> Result<AnimationInfo> {
    if record.record_type != PptRecordType::AnimationInfo {
        return Err(PptError::InvalidFormat(format!(
            "Expected AnimationInfo record, got {:?}",
            record.record_type
        )));
    }

    let mut info = AnimationInfo::new();

    for child in &record.children {
        match child.record_type {
            PptRecordType::BuildList => {
                info.build_list = Some(parse_build_list(child)?);
            },
            PptRecordType::TimeNode => {
                if let Ok(time_node) = parse_time_node(child) {
                    info.time_nodes.push(time_node);
                }
            },
            _ => {
                info.raw_records.push(child.clone());
            },
        }
    }

    Ok(info)
}

/// Parse build list from BuildList container record.
pub fn parse_build_list(record: &PptRecord) -> Result<BuildInfo> {
    if record.record_type != PptRecordType::BuildList {
        return Err(PptError::InvalidFormat(format!(
            "Expected BuildList record, got {:?}",
            record.record_type
        )));
    }

    let mut build_info = BuildInfo::new();

    for child in &record.children {
        match child.record_type {
            PptRecordType::BuildAtom => {
                if let Ok(build) = parse_build_atom(child) {
                    build_info.add_build(build);
                }
            },
            PptRecordType::ChartBuild | PptRecordType::DiagramBuild | PptRecordType::ParaBuild => {
                if let Ok(build) = parse_complex_build(child) {
                    build_info.add_build(build);
                }
            },
            _ => {},
        }
    }

    Ok(build_info)
}

/// Parse a single BuildAtom record.
fn parse_build_atom(record: &PptRecord) -> Result<BuildLevel> {
    if record.data.len() < 16 {
        return Err(PptError::Corrupted(
            "BuildAtom record too small".to_string(),
        ));
    }

    let shape_id = U32::<LittleEndian>::read_from_bytes(&record.data[0..4])
        .map(|v| v.get())
        .unwrap_or(0);

    let build_order = U32::<LittleEndian>::read_from_bytes(&record.data[4..8])
        .map(|v| v.get())
        .unwrap_or(0);

    let flags = U32::<LittleEndian>::read_from_bytes(&record.data[8..12])
        .map(|v| v.get())
        .unwrap_or(0);

    let effect_type = U32::<LittleEndian>::read_from_bytes(&record.data[12..16])
        .map(|v| v.get())
        .unwrap_or(0);

    let build_type = parse_build_type(flags);
    let effect = parse_effect_type(effect_type);
    let speed = parse_effect_speed(flags);
    let direction = parse_effect_direction(flags);
    let trigger = parse_animation_trigger(flags);
    let after_effect = parse_after_effect(flags);
    let iteration = parse_iteration_type(flags);

    Ok(BuildLevel {
        build_type,
        shape_id,
        build_order,
        effect,
        speed,
        direction,
        trigger,
        motion_path: None,
        sound: None,
        iteration,
        after_effect,
        duration_ms: None,
    })
}

/// Parse complex build types (chart, diagram, paragraph).
fn parse_complex_build(record: &PptRecord) -> Result<BuildLevel> {
    let mut build = BuildLevel::default();

    if record.data.len() >= 4 {
        build.shape_id = U32::<LittleEndian>::read_from_bytes(&record.data[0..4])
            .map(|v| v.get())
            .unwrap_or(0);
    }

    build.build_type = match record.record_type {
        PptRecordType::ChartBuild => BuildType::Entrance,
        PptRecordType::DiagramBuild => BuildType::Entrance,
        PptRecordType::ParaBuild => BuildType::Entrance,
        _ => BuildType::Entrance,
    };

    Ok(build)
}

/// Parse time node container.
fn parse_time_node(record: &PptRecord) -> Result<TimeNodeContainer> {
    let mut node = TimeNodeContainer {
        raw_record: Some(record.clone()),
        ..Default::default()
    };

    for child in &record.children {
        match child.record_type {
            PptRecordType::TimePropertyList => {
                if let Ok((duration, delay, fill, restart)) = parse_time_properties(child) {
                    node.duration = duration;
                    node.delay = delay;
                    node.fill = fill;
                    node.restart = restart;
                }
            },
            PptRecordType::TimeBehavior => {},
            PptRecordType::TimeNode => {
                if let Ok(child_node) = parse_time_node(child) {
                    node.children.push(child_node);
                }
            },
            _ => {},
        }
    }

    node.node_type = infer_node_type(&node);

    Ok(node)
}

/// Parse time properties from TimePropertyList record.
fn parse_time_properties(record: &PptRecord) -> Result<(Option<u32>, u32, FillMode, RestartMode)> {
    let mut duration = None;
    let mut delay = 0;
    let fill = FillMode::Hold;
    let restart = RestartMode::Never;

    if record.data.len() >= 8 {
        duration = Some(
            U32::<LittleEndian>::read_from_bytes(&record.data[0..4])
                .map(|v| v.get())
                .unwrap_or(1000),
        );
        delay = U32::<LittleEndian>::read_from_bytes(&record.data[4..8])
            .map(|v| v.get())
            .unwrap_or(0);
    }

    Ok((duration, delay, fill, restart))
}

/// Infer time node type from its structure.
fn infer_node_type(node: &TimeNodeContainer) -> TimeNodeType {
    if node.children.is_empty() {
        TimeNodeType::Effect
    } else {
        TimeNodeType::Sequence
    }
}

/// Parse build type from flags.
fn parse_build_type(flags: u32) -> BuildType {
    let build_type_bits = (flags >> 4) & 0x03;
    match build_type_bits {
        0 => BuildType::Entrance,
        1 => BuildType::Emphasis,
        2 => BuildType::Exit,
        3 => BuildType::MotionPath,
        _ => BuildType::Entrance,
    }
}

/// Parse animation effect type.
fn parse_effect_type(effect_type: u32) -> AnimationEffect {
    match effect_type {
        0 => AnimationEffect::Appear,
        1 => AnimationEffect::FlyIn,
        2 => AnimationEffect::Blinds,
        3 => AnimationEffect::Box,
        4 => AnimationEffect::Checkerboard,
        5 => AnimationEffect::Dissolve,
        6 => AnimationEffect::Split,
        7 => AnimationEffect::Wipe,
        8 => AnimationEffect::RandomBars,
        9 => AnimationEffect::FadeIn,
        10 => AnimationEffect::Zoom,
        11 => AnimationEffect::Swivel,
        12 => AnimationEffect::Bounce,
        13 => AnimationEffect::Pulse,
        14 => AnimationEffect::Spin,
        15 => AnimationEffect::GrowAndTurn,
        16 => AnimationEffect::Teeter,
        17 => AnimationEffect::Wave,
        _ => AnimationEffect::Custom,
    }
}

/// Parse effect speed from flags.
fn parse_effect_speed(flags: u32) -> EffectSpeed {
    let speed_bits = (flags >> 16) & 0x07;
    match speed_bits {
        0 => EffectSpeed::VerySlow,
        1 => EffectSpeed::Slow,
        2 => EffectSpeed::Medium,
        3 => EffectSpeed::Fast,
        4 => EffectSpeed::VeryFast,
        _ => EffectSpeed::Medium,
    }
}

/// Parse effect direction from flags.
fn parse_effect_direction(flags: u32) -> EffectDirection {
    let direction_bits = (flags >> 20) & 0x0F;
    match direction_bits {
        0 => EffectDirection::None,
        1 => EffectDirection::FromTop,
        2 => EffectDirection::FromBottom,
        3 => EffectDirection::FromLeft,
        4 => EffectDirection::FromRight,
        5 => EffectDirection::FromTopLeft,
        6 => EffectDirection::FromTopRight,
        7 => EffectDirection::FromBottomLeft,
        8 => EffectDirection::FromBottomRight,
        _ => EffectDirection::None,
    }
}

/// Parse animation trigger from flags.
fn parse_animation_trigger(flags: u32) -> AnimationTrigger {
    let trigger_bits = flags & 0x03;
    match trigger_bits {
        0 => AnimationTrigger::OnClick,
        1 => AnimationTrigger::WithPrevious,
        2 => AnimationTrigger::AfterPrevious,
        _ => AnimationTrigger::OnClick,
    }
}

/// Parse after-effect from flags.
fn parse_after_effect(flags: u32) -> AfterEffect {
    let after_bits = (flags >> 24) & 0x03;
    match after_bits {
        0 => AfterEffect::None,
        1 => AfterEffect::DimToColor,
        2 => AfterEffect::Hide,
        3 => AfterEffect::HideOnNextClick,
        _ => AfterEffect::None,
    }
}

/// Parse iteration type from flags.
fn parse_iteration_type(flags: u32) -> IterationType {
    let iter_bits = (flags >> 26) & 0x03;
    match iter_bits {
        0 => IterationType::All,
        1 => IterationType::ByElement,
        2 => IterationType::ByWord,
        3 => IterationType::ByLetter,
        _ => IterationType::All,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_build_type() {
        assert_eq!(parse_build_type(0x00), BuildType::Entrance);
        assert_eq!(parse_build_type(0x10), BuildType::Emphasis);
        assert_eq!(parse_build_type(0x20), BuildType::Exit);
        assert_eq!(parse_build_type(0x30), BuildType::MotionPath);
    }

    #[test]
    fn test_parse_effect_speed() {
        assert_eq!(parse_effect_speed(0x000000), EffectSpeed::VerySlow);
        assert_eq!(parse_effect_speed(0x010000), EffectSpeed::Slow);
        assert_eq!(parse_effect_speed(0x020000), EffectSpeed::Medium);
        assert_eq!(parse_effect_speed(0x030000), EffectSpeed::Fast);
        assert_eq!(parse_effect_speed(0x040000), EffectSpeed::VeryFast);
    }

    #[test]
    fn test_parse_animation_trigger() {
        assert_eq!(parse_animation_trigger(0x00), AnimationTrigger::OnClick);
        assert_eq!(
            parse_animation_trigger(0x01),
            AnimationTrigger::WithPrevious
        );
        assert_eq!(
            parse_animation_trigger(0x02),
            AnimationTrigger::AfterPrevious
        );
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
