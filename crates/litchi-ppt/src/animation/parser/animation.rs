//! `PowerPoint` 97 animation-info records.

use super::support::{read_i16, read_i32, read_u16, read_u32};
use crate::animation::triggers::IterationType;
use crate::animation::types::{
    AfterEffect, AnimationInfo, LegacyAnimationAtom, LegacyAnimationBuild, LegacyAnimationEffect,
    LegacyTextBuildSubEffect,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

/// Parse animation info from `AnimationInfo` container record.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_animation_info(record: &Record) -> Result<AnimationInfo> {
    if record.record_type != RecordType::AnimationInfo {
        return Err(Error::InvalidFormat(format!(
            "Expected AnimationInfo record, got {:?}",
            record.record_type
        )));
    }
    if record.version != 0x0F || record.instance != 0 {
        return Err(Error::Corrupted(format!(
            "AnimationInfo requires version 15 and instance 0; got version {} and instance {}",
            record.version, record.instance
        )));
    }

    let mut info = AnimationInfo::new();
    let atom_record = record.children.first().ok_or_else(|| {
        Error::Corrupted("AnimationInfo is missing its AnimationInfoAtom".to_string())
    })?;
    if atom_record.record_type != RecordType::AnimationInfoAtom {
        return Err(Error::InvalidFormat(
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
        if child.record_type == RecordType::AnimationInfoAtom {
            return Err(Error::InvalidFormat(
                "AnimationInfo contains multiple AnimationInfoAtom records".to_string(),
            ));
        }
        info.raw_records.push(child.clone());
    }

    Ok(info)
}

/// Parse the exact `PowerPoint` 97 `AnimationInfoAtom` payload.
///
/// # Errors
///
/// Returns an error if the record is not a 28-byte version-1 `AnimationInfoAtom`,
/// or if any field holds a value the format does not permit.
pub fn parse_animation_info_atom(record: &Record) -> Result<LegacyAnimationAtom> {
    if record.record_type != RecordType::AnimationInfoAtom {
        return Err(Error::InvalidFormat(format!(
            "Expected AnimationInfoAtom record, got {:?}",
            record.record_type
        )));
    }
    if record.version != 1 || record.instance != 0 || record.data.len() != 28 {
        return Err(Error::Corrupted(format!(
            "AnimationInfoAtom requires version 1, instance 0, and 28 data bytes; got version {}, instance {}, length {}",
            record.version,
            record.instance,
            record.data.len()
        )));
    }

    let dim_color = read_u32(&record.data, 0);
    let flags = read_u16(&record.data, 4);
    let mut decoded_flags = [false; 8];
    for (index, decoded) in decoded_flags.iter_mut().enumerate() {
        let value = (flags >> (index * 2)) & 0x03;
        if value > 1 {
            return Err(Error::InvalidFormat(format!(
                "AnimationInfoAtom flag field {index} has invalid bool2 value {value}"
            )));
        }
        *decoded = value == 1;
    }
    let sound_id_ref = read_u32(&record.data, 8);
    let delay_time_ms = read_i32(&record.data, 12);
    if decoded_flags[1] && delay_time_ms < 0 {
        return Err(Error::InvalidFormat(
            "automatic AnimationInfoAtom has a negative delay".to_string(),
        ));
    }
    let order_id = read_i16(&record.data, 16);
    if order_id < -2 {
        return Err(Error::InvalidFormat(format!(
            "AnimationInfoAtom orderID {order_id} is less than -2"
        )));
    }
    let slide_count = read_u16(&record.data, 18);
    let build_type = LegacyAnimationBuild::parse(record.data[20]).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "invalid AnimationInfoAtom animBuildType {:#04X}",
            record.data[20]
        ))
    })?;
    let effect = LegacyAnimationEffect::parse(record.data[21]).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "invalid AnimationInfoAtom animEffect {:#04X}",
            record.data[21]
        ))
    })?;
    let effect_direction = record.data[22];
    if !effect.accepts_direction(effect_direction) {
        return Err(Error::InvalidFormat(format!(
            "AnimationInfoAtom direction {effect_direction:#04X} is invalid for {effect:?}"
        )));
    }
    let after_effect = match record.data[23] {
        0 => AfterEffect::None,
        1 => AfterEffect::DimToColor,
        2 => AfterEffect::HideOnNextClick,
        3 => AfterEffect::Hide,
        value => {
            return Err(Error::InvalidFormat(format!(
                "invalid AnimationInfoAtom animAfterEffect {value:#04X}"
            )));
        },
    };
    let text_build_sub_effect =
        LegacyTextBuildSubEffect::parse(record.data[24]).ok_or_else(|| {
            Error::InvalidFormat(format!(
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
