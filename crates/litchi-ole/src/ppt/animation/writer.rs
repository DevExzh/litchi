//! Animation record writer.
//!
//! Writes PowerPoint binary animation records from structured types.

use super::types::{
    AfterEffect, AnimationEffect, AnimationInfo, AnimationTrigger, BuildInfo, BuildLevel,
    BuildType, EffectDirection, EffectSpeed, LegacyAnimationAtom, LegacyAnimationBuild,
    LegacyAnimationEffect, LegacyTextBuildSubEffect,
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

/// Map animation effect to PPT97 fly method and direction.
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
pub fn write_build_list(build_info: &BuildInfo) -> Vec<u8> {
    let mut data = Vec::new();

    let mut children: Vec<u8> = Vec::new();

    for build in &build_info.builds {
        children.extend(write_build_atom(build));
    }

    let header = create_record_header(PptRecordType::BuildList, 0x0F, 0, children.len() as u32);
    data.extend(header);
    data.extend(children);

    data
}

/// Write BuildAtom record.
fn write_build_atom(build: &BuildLevel) -> Vec<u8> {
    let mut data = Vec::new();

    let mut atom_data: Vec<u8> = Vec::new();
    atom_data.extend(&build.shape_id.to_le_bytes());
    atom_data.extend(&build.build_order.to_le_bytes());

    let flags = encode_build_flags(build);
    atom_data.extend(&flags.to_le_bytes());

    let effect_type = encode_effect_type(build.effect);
    atom_data.extend(&effect_type.to_le_bytes());

    let header = create_record_header(PptRecordType::BuildAtom, 0x01, 0, atom_data.len() as u32);
    data.extend(header);
    data.extend(atom_data);

    data
}

/// Encode build flags.
fn encode_build_flags(build: &BuildLevel) -> u32 {
    let mut flags = 0u32;

    flags |= encode_animation_trigger(build.trigger);

    flags |= (encode_build_type(build.build_type) as u32) << 4;

    flags |= (encode_effect_speed(build.speed) as u32) << 16;

    flags |= (encode_effect_direction(build.direction) as u32) << 20;

    flags |= (encode_after_effect(build.after_effect) as u32) << 24;

    flags |= (encode_iteration_type(&build.iteration) as u32) << 26;

    flags
}

/// Encode build type.
fn encode_build_type(build_type: BuildType) -> u8 {
    match build_type {
        BuildType::Entrance => 0,
        BuildType::Emphasis => 1,
        BuildType::Exit => 2,
        BuildType::MotionPath => 3,
    }
}

/// Encode animation effect type.
fn encode_effect_type(effect: AnimationEffect) -> u32 {
    use AnimationEffect::*;
    match effect {
        // Map core effects to their codes
        Appear => 0,
        FlyIn | FloatIn | Ascend | CrawlIn | RiseUp => 1,
        Blinds | BlindsOut => 2,
        Box | BoxOut => 3,
        Checkerboard | CheckerboardOut => 4,
        Dissolve => 5,
        Split | SplitOut => 6,
        Wipe | WipeOut => 7,
        RandomBars | RandomBarsOut => 8,
        FadeIn | FadeOut => 9,
        Zoom => 10,
        Swivel => 11,
        Bounce => 12,
        Pulse | ColorPulse => 13,
        Spin => 14,
        GrowAndTurn => 15,
        Teeter => 16,
        Wave => 17,
        // New entrance effects map to closest equivalents
        Descend | DescendOut | SinkDown => 1,
        Expand | Stretch | GrowShrink => 10,
        Compress | Collapse => 10,
        Wheel | Strips | StripsOut | PeekIn | PeekOut => 0,
        Plus | PlusOut | Diamond | DiamondOut | Wedge => 0,
        Random | Disappear => 0,
        SpiralIn | SpiralOut => 0,
        // Emphasis effects
        Lighten | Darken | ChangeFillColor | ChangeLineColor => 0,
        ChangeFontColor | ChangeFontSize | BoldFlash | Underline => 0,
        ComplementaryColor | ComplementaryColor2 | ContrastingColor => 0,
        Transparency | ObjectColor | VerticalHighlight | Flicker => 0,
        // Exit effects
        FlyOut | CrawlOut => 1,
        // Motion paths (not supported, map to 0)
        MotionPath | MotionPathLines | MotionPathCurves | MotionPathShapes => 0,
        MotionPathLeft | MotionPathRight | MotionPathUp | MotionPathDown => 0,
        MotionPathDiagonalUpRight | MotionPathDiagonalDownRight => 0,
        MotionPathArcDown | MotionPathArcUp | MotionPathCircle => 0,
        MotionPathDiamond | MotionPathHeart | MotionPathHexagon => 0,
        MotionPathOctagon | MotionPathPentagon | MotionPathSquare => 0,
        MotionPathStar4 | MotionPathStar5 | MotionPathStar6 | MotionPathStar8 => 0,
        MotionPathTriangle | MotionPathLoopDeLoop | MotionPathCurvedX => 0,
        MotionPathSCurve1 | MotionPathSCurve2 | MotionPathSineWave => 0,
        MotionPathSpiralLeft | MotionPathSpiralRight | MotionPathSpring => 0,
        MotionPathZigzag => 0,
        Custom => 255,
    }
}

/// Encode effect speed.
fn encode_effect_speed(speed: EffectSpeed) -> u8 {
    match speed {
        EffectSpeed::VerySlow => 0,
        EffectSpeed::Slow => 1,
        EffectSpeed::Medium => 2,
        EffectSpeed::Fast => 3,
        EffectSpeed::VeryFast => 4,
    }
}

/// Encode effect direction.
fn encode_effect_direction(direction: EffectDirection) -> u8 {
    use EffectDirection::*;
    match direction {
        None => 0,
        FromTop => 1,
        FromBottom => 2,
        FromLeft => 3,
        FromRight => 4,
        FromTopLeft => 5,
        FromTopRight => 6,
        FromBottomLeft => 7,
        FromBottomRight => 8,
        Horizontal => 0,
        Vertical => 1,
        In => 0,
        Out => 1,
        Across => 0,
        Clockwise => 0,
        CounterClockwise => 1,
    }
}

/// Encode animation trigger.
fn encode_animation_trigger(trigger: AnimationTrigger) -> u32 {
    match trigger {
        AnimationTrigger::OnClick => 0,
        AnimationTrigger::WithPrevious => 1,
        AnimationTrigger::AfterPrevious => 2,
    }
}

/// Encode after-effect.
fn encode_after_effect(after_effect: AfterEffect) -> u8 {
    match after_effect {
        AfterEffect::None => 0,
        AfterEffect::DimToColor => 1,
        AfterEffect::HideOnNextClick => 2,
        AfterEffect::Hide => 3,
    }
}

/// Encode iteration type.
fn encode_iteration_type(iteration: &super::triggers::IterationType) -> u8 {
    use super::triggers::IterationType;
    match iteration {
        IterationType::All => 0,
        IterationType::ByElement => 1,
        IterationType::ByWord => 2,
        IterationType::ByLetter => 3,
    }
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

    #[test]
    fn test_encode_build_type() {
        assert_eq!(encode_build_type(BuildType::Entrance), 0);
        assert_eq!(encode_build_type(BuildType::Emphasis), 1);
        assert_eq!(encode_build_type(BuildType::Exit), 2);
        assert_eq!(encode_build_type(BuildType::MotionPath), 3);
    }

    #[test]
    fn test_encode_effect_speed() {
        assert_eq!(encode_effect_speed(EffectSpeed::VerySlow), 0);
        assert_eq!(encode_effect_speed(EffectSpeed::Slow), 1);
        assert_eq!(encode_effect_speed(EffectSpeed::Medium), 2);
        assert_eq!(encode_effect_speed(EffectSpeed::Fast), 3);
        assert_eq!(encode_effect_speed(EffectSpeed::VeryFast), 4);
    }

    #[test]
    fn test_encode_animation_trigger() {
        assert_eq!(encode_animation_trigger(AnimationTrigger::OnClick), 0);
        assert_eq!(encode_animation_trigger(AnimationTrigger::WithPrevious), 1);
        assert_eq!(encode_animation_trigger(AnimationTrigger::AfterPrevious), 2);
    }

    #[test]
    fn test_write_build_atom_header() {
        let build = BuildLevel::default();
        let data = write_build_atom(&build);

        assert!(data.len() >= 8);
    }

    #[test]
    fn test_write_build_list_empty() {
        let build_info = BuildInfo::new();
        let data = write_build_list(&build_info);

        assert_eq!(data.len(), 8);
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
