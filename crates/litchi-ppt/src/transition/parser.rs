//! Slide transition parser.
//!
//! Parses `PowerPoint` binary slide transition records.

use super::types::{
    AdvanceMode, SoundAction, TransitionDirection, TransitionInfo, TransitionSpeed, TransitionType,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;
use zerocopy::{FromBytes, byteorder::LittleEndian, byteorder::U16, byteorder::U32};

/// Parse transition info from `SSSlideInfoAtom` record.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_transition(record: &Record) -> Result<TransitionInfo> {
    if record.record_type != RecordType::SSSlideInfoAtom {
        return Err(Error::InvalidFormat(format!(
            "Expected SSSlideInfoAtom record, got {:?}",
            record.record_type
        )));
    }

    if record.data.len() < 16 {
        return Err(Error::Corrupted(
            "SSSlideInfoAtom record too small".to_string(),
        ));
    }

    let mut transition = TransitionInfo::new();

    let slide_time = U32::<LittleEndian>::read_from_bytes(&record.data[0..4]).map_or(0, U32::get);

    if slide_time > 0 {
        transition.advance_time_ms = Some(slide_time);
    }

    let sound_id_ref = U32::<LittleEndian>::read_from_bytes(&record.data[4..8]).map_or(0, U32::get);

    let effect_direction = record.data.get(8).copied().unwrap_or(0);

    let effect_type = record.data.get(9).copied().map_or(0, u16::from);

    let flags = U16::<LittleEndian>::read_from_bytes(&record.data[10..12])
        .map_or(0, |v| u32::from(v.get()));

    let effect_speed = record.data.get(12).copied().unwrap_or(0);

    (transition.transition_type, transition.direction) =
        parse_transition_visual(effect_type, effect_direction);
    transition.speed = parse_transition_speed(effect_speed);
    transition.advance_mode = parse_advance_mode(flags, slide_time > 0);
    transition.loop_sound = (flags & 0x40) != 0;

    if sound_id_ref > 0 {
        transition.sound = Some(parse_sound_action(sound_id_ref));
    }

    Ok(transition)
}

/// Parse one transition kind/direction pair using the exact effect table in
/// `[MS-PPT]` 2.6.6. Invalid or undefined producer values remain explicitly
/// undefined instead of being mislabeled as another effect.
pub(super) fn parse_transition_visual(
    effect_type: u16,
    direction: u8,
) -> (TransitionType, TransitionDirection) {
    let undefined = (TransitionType::Undefined, TransitionDirection::None);
    match effect_type {
        0 => match direction {
            0 => (TransitionType::None, TransitionDirection::None),
            1 => (TransitionType::Cut, TransitionDirection::ThroughBlack),
            _ => undefined,
        },
        1 => (TransitionType::Random, TransitionDirection::None),
        2 => match direction {
            0 => (TransitionType::Blinds, TransitionDirection::Vertical),
            1 => (TransitionType::Blinds, TransitionDirection::Horizontal),
            _ => undefined,
        },
        3 | 8 | 21 => match direction {
            0 => (
                match effect_type {
                    3 => TransitionType::Checkerboard,
                    8 => TransitionType::RandomBars,
                    _ => TransitionType::Comb,
                },
                TransitionDirection::Horizontal,
            ),
            1 => (
                match effect_type {
                    3 => TransitionType::Checkerboard,
                    8 => TransitionType::RandomBars,
                    _ => TransitionType::Comb,
                },
                TransitionDirection::Vertical,
            ),
            _ => undefined,
        },
        4 | 7 => {
            let transition_type = if effect_type == 4 {
                TransitionType::Cover
            } else {
                TransitionType::Uncover
            };
            let transition_direction = match direction {
                0 => TransitionDirection::FromLeft,
                1 => TransitionDirection::FromTop,
                2 => TransitionDirection::FromRight,
                3 => TransitionDirection::FromBottom,
                4 => TransitionDirection::LeftUp,
                5 => TransitionDirection::RightUp,
                6 => TransitionDirection::LeftDown,
                7 => TransitionDirection::RightDown,
                _ => return undefined,
            };
            (transition_type, transition_direction)
        },
        5 | 6 | 17 | 18 | 19 | 22 | 23 | 27 if direction == 0 => (
            match effect_type {
                5 => TransitionType::Dissolve,
                6 => TransitionType::Fade,
                17 => TransitionType::Diamond,
                18 => TransitionType::Plus,
                19 => TransitionType::Wedge,
                22 => TransitionType::Newsflash,
                23 => TransitionType::AlphaFade,
                _ => TransitionType::Circle,
            },
            TransitionDirection::None,
        ),
        9 => match direction {
            4 => (TransitionType::Strips, TransitionDirection::LeftUp),
            5 => (TransitionType::Strips, TransitionDirection::RightUp),
            6 => (TransitionType::Strips, TransitionDirection::LeftDown),
            7 => (TransitionType::Strips, TransitionDirection::RightDown),
            _ => undefined,
        },
        10 | 20 => {
            let transition_type = if effect_type == 10 {
                TransitionType::Wipe
            } else {
                TransitionType::Push
            };
            let transition_direction = match direction {
                0 => TransitionDirection::FromLeft,
                1 => TransitionDirection::FromTop,
                2 => TransitionDirection::FromRight,
                3 => TransitionDirection::FromBottom,
                _ => return undefined,
            };
            (transition_type, transition_direction)
        },
        11 => match direction {
            0 => (TransitionType::Box, TransitionDirection::Out),
            1 => (TransitionType::Box, TransitionDirection::In),
            _ => undefined,
        },
        13 => match direction {
            0 => (TransitionType::Split, TransitionDirection::HorizontalOut),
            1 => (TransitionType::Split, TransitionDirection::HorizontalIn),
            2 => (TransitionType::Split, TransitionDirection::VerticalOut),
            3 => (TransitionType::Split, TransitionDirection::VerticalIn),
            _ => undefined,
        },
        26 => match direction {
            1 => (TransitionType::Wheel, TransitionDirection::Spokes1),
            2 => (TransitionType::Wheel, TransitionDirection::Spokes2),
            3 => (TransitionType::Wheel, TransitionDirection::Spokes3),
            4 => (TransitionType::Wheel, TransitionDirection::Spokes4),
            8 => (TransitionType::Wheel, TransitionDirection::Spokes8),
            _ => undefined,
        },
        _ => undefined,
    }
}

/// Parse transition speed from speed value.
pub(super) fn parse_transition_speed(speed: u8) -> TransitionSpeed {
    match speed {
        0 => TransitionSpeed::Slow,
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

    SoundAction::builtin(format!("Sound{sound_id}"))
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    #[test]
    fn parses_spec_transition_visuals_without_shifted_names() {
        assert_eq!(
            parse_transition_visual(0, 0),
            (TransitionType::None, TransitionDirection::None)
        );
        assert_eq!(
            parse_transition_visual(1, 99),
            (TransitionType::Random, TransitionDirection::None)
        );
        assert_eq!(
            parse_transition_visual(4, 3),
            (TransitionType::Cover, TransitionDirection::FromBottom)
        );
        assert_eq!(
            parse_transition_visual(11, 1),
            (TransitionType::Box, TransitionDirection::In)
        );
        assert_eq!(
            parse_transition_visual(17, 0),
            (TransitionType::Diamond, TransitionDirection::None)
        );
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

    fn slide_info_record(data: Vec<u8>) -> Record {
        Record {
            record_type: RecordType::SSSlideInfoAtom,
            record_type_raw: 0x03F9,
            version: 0,
            instance: 0,
            data_length: u32::try_from(data.len()).unwrap(),
            data,
            children: Vec::new(),
        }
    }

    #[test]
    fn rejects_truncated_slide_show_slide_info_atom() {
        let record = slide_info_record(vec![0u8; 8]);
        assert!(parse_transition(&record).is_err());
    }

    #[test]
    fn rejects_wrong_record_type() {
        let mut record = slide_info_record(vec![0u8; 16]);
        record.record_type = RecordType::CString;
        assert!(parse_transition(&record).is_err());
    }

    #[test]
    fn parses_minimal_slide_show_slide_info_atom() {
        // slideTime=2000ms, soundIdRef=0, direction=0, effect=10 (Wipe),
        // flags=auto-advance bit only, speed=2 (Fast)
        let mut data = vec![0u8; 16];
        data[0..4].copy_from_slice(&2000u32.to_le_bytes());
        data[9] = 10;
        data[10..12].copy_from_slice(&0x0400u16.to_le_bytes());
        data[12] = 2;

        let info = parse_transition(&slide_info_record(data)).expect("parse transition");
        assert_eq!(info.transition_type, TransitionType::Wipe);
        assert_eq!(info.speed, TransitionSpeed::Fast);
        assert_eq!(info.advance_mode, AdvanceMode::Automatic);
        assert_eq!(info.advance_time_ms, Some(2000));
        assert!(info.sound.is_none());
    }
}
