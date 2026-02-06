//! Slide transition parser.
//!
//! Parses PowerPoint binary slide transition records.

use super::types::{
    AdvanceMode, SoundAction, TransitionDirection, TransitionInfo, TransitionSpeed, TransitionType,
};
use crate::ole::consts::PptRecordType;
use crate::ole::ppt::package::{PptError, Result};
use crate::ole::ppt::records::PptRecord;
use zerocopy::{FromBytes, byteorder::LittleEndian, byteorder::U16, byteorder::U32};

/// Parse transition info from SSSlideInfoAtom record.
pub fn parse_transition(record: &PptRecord) -> Result<TransitionInfo> {
    if record.record_type != PptRecordType::SSSlideInfoAtom {
        return Err(PptError::InvalidFormat(format!(
            "Expected SSSlideInfoAtom record, got {:?}",
            record.record_type
        )));
    }

    if record.data.len() < 20 {
        return Err(PptError::Corrupted(
            "SSSlideInfoAtom record too small".to_string(),
        ));
    }

    let mut transition = TransitionInfo::new();

    let slide_time = U32::<LittleEndian>::read_from_bytes(&record.data[0..4])
        .map(|v| v.get())
        .unwrap_or(0);

    if slide_time > 0 {
        transition.advance_time_ms = Some(slide_time);
    }

    let sound_id_ref = U32::<LittleEndian>::read_from_bytes(&record.data[4..8])
        .map(|v| v.get())
        .unwrap_or(0);

    let effect_type = U16::<LittleEndian>::read_from_bytes(&record.data[8..10])
        .map(|v| v.get())
        .unwrap_or(0);

    let effect_direction = record.data.get(10).copied().unwrap_or(0);

    let effect_speed = record.data.get(11).copied().unwrap_or(0);

    let flags = U32::<LittleEndian>::read_from_bytes(&record.data[12..16])
        .map(|v| v.get())
        .unwrap_or(0);

    transition.transition_type = parse_transition_type(effect_type);
    transition.direction = parse_transition_direction(effect_direction, effect_type);
    transition.speed = parse_transition_speed(effect_speed);
    transition.advance_mode = parse_advance_mode(flags, slide_time > 0);
    transition.loop_sound = (flags & 0x40) != 0;

    if sound_id_ref > 0 {
        transition.sound = Some(parse_sound_action(sound_id_ref));
    }

    Ok(transition)
}

/// Parse transition type from effect type value.
fn parse_transition_type(effect_type: u16) -> TransitionType {
    match effect_type {
        0 => TransitionType::None,
        1 => TransitionType::Blinds,
        2 => TransitionType::Checkerboard,
        3 => TransitionType::Cover,
        4 => TransitionType::Dissolve,
        5 => TransitionType::Fade,
        6 => TransitionType::Uncover,
        7 => TransitionType::RandomBars,
        8 => TransitionType::Strips,
        9 => TransitionType::Wipe,
        10 => TransitionType::Box,
        11 => TransitionType::Random,
        13 => TransitionType::Split,
        17 => TransitionType::Cut,
        18 => TransitionType::Push,
        19 => TransitionType::Comb,
        20 => TransitionType::Zoom,
        21 => TransitionType::Wedge,
        22 => TransitionType::Wheel,
        23 => TransitionType::Newsflash,
        24 => TransitionType::Vortex,
        25 => TransitionType::Shred,
        26 => TransitionType::Switch,
        27 => TransitionType::Flip,
        28 => TransitionType::Gallery,
        29 => TransitionType::Cube,
        30 => TransitionType::Doors,
        31 => TransitionType::Window,
        32 => TransitionType::Ferris,
        33 => TransitionType::Conveyor,
        34 => TransitionType::Rotate,
        35 => TransitionType::Pan,
        36 => TransitionType::Glitter,
        37 => TransitionType::Honeycomb,
        38 => TransitionType::Flash,
        39 => TransitionType::Ripple,
        40 => TransitionType::Fracture,
        41 => TransitionType::Crush,
        42 => TransitionType::Peel,
        43 => TransitionType::PageCurl,
        44 => TransitionType::Airplane,
        45 => TransitionType::Origami,
        46 => TransitionType::Morph,
        _ => TransitionType::None,
    }
}

/// Parse transition direction from effect direction value.
fn parse_transition_direction(direction: u8, effect_type: u16) -> TransitionDirection {
    match effect_type {
        1 => {
            if direction == 0 {
                TransitionDirection::Horizontal
            } else {
                TransitionDirection::Vertical
            }
        },
        2 => {
            if direction == 0 {
                TransitionDirection::Horizontal
            } else {
                TransitionDirection::Vertical
            }
        },
        3 | 6 | 9 => match direction {
            0 => TransitionDirection::FromLeft,
            1 => TransitionDirection::FromTop,
            2 => TransitionDirection::FromRight,
            3 => TransitionDirection::FromBottom,
            _ => TransitionDirection::None,
        },
        7 => {
            if direction == 0 {
                TransitionDirection::Horizontal
            } else {
                TransitionDirection::Vertical
            }
        },
        8 => match direction {
            0 => TransitionDirection::LeftDown,
            1 => TransitionDirection::LeftUp,
            2 => TransitionDirection::RightDown,
            3 => TransitionDirection::RightUp,
            _ => TransitionDirection::None,
        },
        10 => {
            if direction == 0 {
                TransitionDirection::In
            } else {
                TransitionDirection::Out
            }
        },
        13 => match direction {
            0 => TransitionDirection::Horizontal,
            1 => TransitionDirection::Vertical,
            _ => TransitionDirection::None,
        },
        _ => TransitionDirection::None,
    }
}

/// Parse transition speed from speed value.
fn parse_transition_speed(speed: u8) -> TransitionSpeed {
    match speed {
        0 => TransitionSpeed::Slow,
        1 => TransitionSpeed::Medium,
        2 => TransitionSpeed::Fast,
        _ => TransitionSpeed::Medium,
    }
}

/// Parse advance mode from flags and timing.
fn parse_advance_mode(flags: u32, has_auto_advance: bool) -> AdvanceMode {
    let advance_on_click = (flags & 0x01) != 0;

    if has_auto_advance && advance_on_click {
        AdvanceMode::Both
    } else if has_auto_advance {
        AdvanceMode::Automatic
    } else {
        AdvanceMode::OnClick
    }
}

/// Parse sound action from sound ID reference.
fn parse_sound_action(sound_id: u32) -> SoundAction {
    let builtin_sounds = [
        (1, "Applause"),
        (2, "Arrow"),
        (3, "Bomb"),
        (4, "Breeze"),
        (5, "Camera"),
        (6, "Cash Register"),
        (7, "Chime"),
        (8, "Click"),
        (9, "Coin"),
        (10, "Drum Roll"),
        (11, "Explosion"),
        (12, "Hammer"),
        (13, "Laser"),
        (14, "Push"),
        (15, "Suction"),
        (16, "Swoosh"),
        (17, "Typewriter"),
        (18, "Voltage"),
        (19, "Whoosh"),
        (20, "Wind"),
    ];

    for (id, name) in &builtin_sounds {
        if *id == sound_id {
            return SoundAction::builtin(*name);
        }
    }

    SoundAction::builtin(format!("Sound{}", sound_id))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_transition_type() {
        assert_eq!(parse_transition_type(0), TransitionType::None);
        assert_eq!(parse_transition_type(1), TransitionType::Blinds);
        assert_eq!(parse_transition_type(4), TransitionType::Dissolve);
        assert_eq!(parse_transition_type(11), TransitionType::Random);
    }

    #[test]
    fn test_parse_transition_speed() {
        assert_eq!(parse_transition_speed(0), TransitionSpeed::Slow);
        assert_eq!(parse_transition_speed(1), TransitionSpeed::Medium);
        assert_eq!(parse_transition_speed(2), TransitionSpeed::Fast);
    }

    #[test]
    fn test_parse_advance_mode() {
        assert_eq!(parse_advance_mode(0x01, false), AdvanceMode::OnClick);
        assert_eq!(parse_advance_mode(0x00, true), AdvanceMode::Automatic);
        assert_eq!(parse_advance_mode(0x01, true), AdvanceMode::Both);
    }

    #[test]
    fn test_transition_info_default() {
        let info = TransitionInfo::default();
        assert_eq!(info.transition_type, TransitionType::None);
        assert_eq!(info.speed, TransitionSpeed::Medium);
        assert!(!info.has_effect());
    }
}
