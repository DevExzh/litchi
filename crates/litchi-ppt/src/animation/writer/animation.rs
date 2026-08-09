//! `PowerPoint` 97 animation-info records.

use super::build::map_effect_to_ppt97;
use super::support::{create_record_header, serialize_raw_record, wrap_record};
use crate::animation::types::{
    AfterEffect, AnimationInfo, LegacyAnimationAtom, LegacyAnimationBuild, LegacyAnimationEffect,
    LegacyTextBuildSubEffect,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};

/// Write `InteractiveInfo` container with `InteractiveInfoAtom` for animations.
/// Per POI `MovieShape`, this is required alongside `AnimationInfo` in `ClientData`.
/// For sound animations, soundRef should match AnimationInfoAtom.soundRef
#[must_use]
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

    #[allow(
        clippy::cast_possible_truncation,
        reason = "the InteractiveInfoAtom payload above is a fixed 16 bytes"
    )]
    let atom_header = create_record_header(
        RecordType::InteractiveInfoAtom,
        0x00,
        0,
        atom_data.len() as u32,
    );

    let mut children = Vec::new();
    children.extend(atom_header);
    children.extend(atom_data);

    // InteractiveInfo container wrapping the atom
    #[allow(
        clippy::cast_possible_truncation,
        reason = "the children are a fixed 8-byte header plus the 16-byte atom"
    )]
    let header = create_record_header(RecordType::InteractiveInfo, 0x0F, 0, children.len() as u32);
    data.extend(header);
    data.extend(children);

    data
}

/// Write `AnimationInfo` container record.
/// Returns (`AnimationInfo` bytes, `sound_ref` for `InteractiveInfo`)
///
/// # Errors
///
/// Returns an error if serialization fails or the underlying writer reports an error.
pub fn write_animation_info(info: &AnimationInfo) -> Result<(Vec<u8>, u32)> {
    if !info.time_nodes.is_empty() {
        return Err(Error::InvalidFormat(
            "extended time nodes belong to the slide animation extension, not AnimationInfo"
                .to_string(),
        ));
    }
    let mut children: Vec<u8> = Vec::new();

    // AnimationInfoAtom MUST be the first child (per POI)
    // Extract first build item to determine animation type and sound
    let atom = if let Some(atom) = &info.legacy_atom {
        atom.clone()
    } else {
        let (fly_method, fly_direction, build_sound) = if let Some(ref build_list) = info.build_list
        {
            if let Some(first_build) = build_list.builds.first() {
                let (method, dir) = map_effect_to_ppt97(first_build.effect, first_build.direction);
                (method, dir, first_build.sound.as_ref())
            } else {
                (0x00, 0, None)
            }
        } else {
            (0x00, 0, None)
        };
        let sound = info.sound.as_ref().or(build_sound);
        LegacyAnimationAtom {
            has_sound: sound.is_some(),
            sound_id_ref: sound.map_or(0, |animation_sound| animation_sound.sound_ref),
            build_type: if info.has_animations() {
                LegacyAnimationBuild::OneBuild
            } else {
                LegacyAnimationBuild::NoBuild
            },
            effect: LegacyAnimationEffect::parse(fly_method).unwrap_or_default(),
            effect_direction: fly_direction,
            text_build_sub_effect: match info.iteration {
                crate::animation::triggers::IterationType::ByWord => {
                    LegacyTextBuildSubEffect::ByWord
                },
                crate::animation::triggers::IterationType::ByLetter => {
                    LegacyTextBuildSubEffect::ByCharacter
                },
                crate::animation::triggers::IterationType::All
                | crate::animation::triggers::IterationType::ByElement => {
                    LegacyTextBuildSubEffect::AllAtOnce
                },
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
        if raw_record.record_type == RecordType::AnimationInfoAtom {
            return Err(Error::InvalidFormat(
                "raw AnimationInfo children cannot contain another AnimationInfoAtom".to_string(),
            ));
        }
        children.extend(serialize_raw_record(raw_record));
    }

    let data = wrap_record(RecordType::AnimationInfo, 0x0F, 0, children)?;

    Ok((data, sound_ref))
}

/// Serialize an exact `PowerPoint` 97 `AnimationInfoAtom`.
///
/// # Errors
///
/// Returns an error if serialization fails or the underlying writer reports an error.
pub fn write_animation_info_atom(atom: &LegacyAnimationAtom) -> Result<Vec<u8>> {
    if atom.automatic && atom.delay_time_ms < 0 {
        return Err(Error::InvalidFormat(
            "automatic AnimationInfoAtom cannot have a negative delay".to_string(),
        ));
    }
    if atom.order_id < -2 {
        return Err(Error::InvalidFormat(format!(
            "AnimationInfoAtom orderID {} is less than -2",
            atom.order_id
        )));
    }
    if !atom.effect.accepts_direction(atom.effect_direction) {
        return Err(Error::InvalidFormat(format!(
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

    let mut result = create_record_header(RecordType::AnimationInfoAtom, 0x01, 0, 28);
    result.extend(data);
    Ok(result)
}
