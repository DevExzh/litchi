//! Animation record parser.
//!
//! Parses PowerPoint binary animation records into structured types.

use super::triggers::IterationType;
use super::types::{
    AfterEffect, AnimationEffect, AnimationInfo, AnimationTrigger, BuildInfo, BuildLevel,
    BuildType, EffectDirection, EffectSpeed, LegacyAnimationAtom, LegacyAnimationBuild,
    LegacyAnimationEffect, LegacyTextBuildSubEffect,
};
use crate::consts::PptRecordType;
use crate::ppt::package::{PptError, Result};
use crate::ppt::records::PptRecord;
use zerocopy::{FromBytes, byteorder::LittleEndian, byteorder::U32};

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
        2 => AfterEffect::HideOnNextClick,
        3 => AfterEffect::Hide,
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
    use crate::ppt::animation::{
        LegacyAnimationAtom, LegacyAnimationBuild, LegacyAnimationEffect, LegacyTextBuildSubEffect,
        write_animation_info, write_animation_info_atom,
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
