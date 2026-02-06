//! Slide transition writer.
//!
//! Writes PowerPoint binary slide transition records.

use super::types::{
    AdvanceMode, TransitionDirection, TransitionInfo, TransitionSpeed, TransitionType,
};
use crate::ole::consts::PptRecordType;

/// Write SSSlideInfoAtom record with transition information.
pub fn write_transition(transition: &TransitionInfo) -> Vec<u8> {
    let mut data = Vec::new();

    // SSSlideInfoAtom structure (16 bytes total):
    // slideTime (4 bytes), soundIdRef (4 bytes), effectDirection (1 byte),
    // effectType (1 byte), effectTransitionFlags (2 bytes), speed (1 byte), unused (3 bytes)

    let slide_time = match transition.advance_mode {
        AdvanceMode::Automatic | AdvanceMode::Both => transition.advance_time_ms.unwrap_or(0),
        AdvanceMode::OnClick => 0,
    };
    data.extend(&slide_time.to_le_bytes());

    let sound_id_ref = if transition.sound.is_some() {
        1u32
    } else {
        0u32
    };
    data.extend(&sound_id_ref.to_le_bytes());

    // effectDirection comes BEFORE effectType (1 byte)
    let effect_direction =
        encode_transition_direction(transition.direction, transition.transition_type);
    data.push(effect_direction);

    // effectType is 1 byte, not 2!
    let effect_type = encode_transition_type(transition.transition_type);
    data.push(effect_type);

    // effectTransitionFlags (2 bytes)
    let flags = encode_transition_flags(transition);
    data.extend(&flags.to_le_bytes());

    // speed (1 byte)
    let effect_speed = encode_transition_speed(transition.speed);
    data.push(effect_speed);

    // unused (3 bytes)
    data.extend(&[0u8, 0u8, 0u8]);

    let header = create_record_header(PptRecordType::SSSlideInfoAtom, 0x00, 0, data.len() as u32);

    let mut result = Vec::new();
    result.extend(header);
    result.extend(data);

    result
}

/// Encode transition type to effect type value (1 byte).
/// Values from LibreOffice pptanimations.hxx
fn encode_transition_type(transition_type: TransitionType) -> u8 {
    match transition_type {
        TransitionType::None => 0,
        TransitionType::Random => 1,
        TransitionType::Blinds => 2,
        TransitionType::Checkerboard => 3, // CHECKER
        TransitionType::Cover => 4,
        TransitionType::Dissolve => 5,
        TransitionType::Fade => 6,
        TransitionType::Uncover => 7, // PULL
        TransitionType::RandomBars => 8,
        TransitionType::Strips => 9,
        TransitionType::Wipe => 10,
        TransitionType::Box => 11,  // ZOOM/Box In/Out
        TransitionType::Zoom => 11, // Same as Box
        TransitionType::Split => 13,
        TransitionType::Cut => 17,       // DIAMOND
        TransitionType::Push => 20,      // Not 18!
        TransitionType::Comb => 21,      // Not 19!
        TransitionType::Newsflash => 22, // Not 23!
        TransitionType::Wedge => 19,     // Not 21!
        TransitionType::Wheel => 26,     // Not 22!
        TransitionType::Vortex => 24,
        TransitionType::Shred => 25,
        TransitionType::Switch => 26,
        TransitionType::Flip => 27,
        TransitionType::Gallery => 28,
        TransitionType::Cube => 29,
        TransitionType::Doors => 30,
        TransitionType::Window => 31,
        TransitionType::Ferris => 32,
        TransitionType::Conveyor => 33,
        TransitionType::Rotate => 34,
        TransitionType::Pan => 35,
        TransitionType::Glitter => 36,
        TransitionType::Honeycomb => 37,
        TransitionType::Flash => 38,
        TransitionType::Ripple => 39,
        TransitionType::Fracture => 40,
        TransitionType::Crush => 41,
        TransitionType::Peel => 42,
        TransitionType::PageCurl => 43,
        TransitionType::Airplane => 44,
        TransitionType::Origami => 45,
        TransitionType::Morph => 46,
    }
}

/// Encode transition direction based on type.
fn encode_transition_direction(
    direction: TransitionDirection,
    transition_type: TransitionType,
) -> u8 {
    match transition_type {
        TransitionType::Blinds => match direction {
            // Per LibreOffice pptin.cxx lines 1578-1583:
            // 0=VERTICAL_STRIPES (vertical blinds), 1=HORIZONTAL_STRIPES
            TransitionDirection::Vertical => 0,
            TransitionDirection::Horizontal => 1,
            _ => 0,
        },
        TransitionType::Checkerboard | TransitionType::RandomBars => match direction {
            // Checkerboard: 0=horizontal, 1=vertical
            TransitionDirection::Horizontal => 0,
            TransitionDirection::Vertical => 1,
            _ => 0,
        },
        TransitionType::Split => match direction {
            // Per LibreOffice pptin.cxx lines 1644-1651:
            // 0=OPEN_VERTICAL (horizontal split out), 1=CLOSE_VERTICAL (horizontal split in)
            // 2=OPEN_HORIZONTAL (vertical split out), 3=CLOSE_HORIZONTAL (vertical split in)
            TransitionDirection::Horizontal => 0, // Horizontal split (opens vertically)
            TransitionDirection::Vertical => 2,   // Vertical split (opens horizontally)
            _ => 0,
        },
        TransitionType::Cover
        | TransitionType::Uncover
        | TransitionType::Wipe
        | TransitionType::Push => match direction {
            // Per LibreOffice pptin.cxx lines 1596-1603:
            // 0=FROM_RIGHT, 1=FROM_BOTTOM, 2=FROM_LEFT, 3=FROM_TOP
            TransitionDirection::FromRight => 0,
            TransitionDirection::FromBottom => 1,
            TransitionDirection::FromLeft => 2,
            TransitionDirection::FromTop => 3,
            _ => 0,
        },
        TransitionType::Strips => match direction {
            TransitionDirection::LeftDown => 0,
            TransitionDirection::LeftUp => 1,
            TransitionDirection::RightDown => 2,
            TransitionDirection::RightUp => 3,
            _ => 0,
        },
        TransitionType::Box | TransitionType::Zoom => match direction {
            // Box/Zoom use same type (11), direction differentiates In vs Out
            TransitionDirection::Out => 0,
            TransitionDirection::In => 1,
            _ => 1,
        },
        _ => 0,
    }
}

/// Encode transition speed.
fn encode_transition_speed(speed: TransitionSpeed) -> u8 {
    match speed {
        TransitionSpeed::Slow => 0,
        TransitionSpeed::Medium => 1,
        TransitionSpeed::Fast => 2,
    }
}

/// Encode transition flags (effectTransitionFlags).
/// Per POI SSSlideInfoAtom:
/// - MANUAL_ADVANCE_BIT = 1 << 0 (bit 0)
/// - HIDDEN_BIT = 1 << 2
/// - SOUND_BIT = 1 << 4
/// - LOOP_SOUND_BIT = 1 << 6
/// - STOP_SOUND_BIT = 1 << 8
/// - AUTO_ADVANCE_BIT = 1 << 10
/// - CURSOR_VISIBLE_BIT = 1 << 12
fn encode_transition_flags(transition: &TransitionInfo) -> u16 {
    let mut flags = 0u16;

    // Manual advance (on click)
    if matches!(
        transition.advance_mode,
        AdvanceMode::OnClick | AdvanceMode::Both
    ) {
        flags |= 1 << 0; // MANUAL_ADVANCE_BIT
    }

    // Auto advance
    if matches!(
        transition.advance_mode,
        AdvanceMode::Automatic | AdvanceMode::Both
    ) {
        flags |= 1 << 10; // AUTO_ADVANCE_BIT
    }

    if transition.loop_sound {
        flags |= 1 << 6; // LOOP_SOUND_BIT
    }

    flags
}

/// Create a PPT record header.
fn create_record_header(
    record_type: PptRecordType,
    version: u16,
    instance: u16,
    data_length: u32,
) -> Vec<u8> {
    let mut header = Vec::with_capacity(8);

    let version_instance = version | (instance << 4);
    header.extend(&version_instance.to_le_bytes());

    header.extend(&record_type.as_u16().to_le_bytes());

    header.extend(&data_length.to_le_bytes());

    header
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_encode_transition_type() {
        assert_eq!(encode_transition_type(TransitionType::None), 0);
        assert_eq!(encode_transition_type(TransitionType::Blinds), 1);
        assert_eq!(encode_transition_type(TransitionType::Dissolve), 4);
        assert_eq!(encode_transition_type(TransitionType::Random), 11);
    }

    #[test]
    fn test_encode_transition_speed() {
        assert_eq!(encode_transition_speed(TransitionSpeed::Slow), 0);
        assert_eq!(encode_transition_speed(TransitionSpeed::Medium), 1);
        assert_eq!(encode_transition_speed(TransitionSpeed::Fast), 2);
    }

    #[test]
    fn test_encode_transition_flags_on_click() {
        let transition = TransitionInfo {
            advance_mode: AdvanceMode::OnClick,
            loop_sound: false,
            ..Default::default()
        };
        let flags = encode_transition_flags(&transition);
        assert_eq!(flags & 0x01, 0x01);
    }

    #[test]
    fn test_encode_transition_flags_loop_sound() {
        let transition = TransitionInfo {
            advance_mode: AdvanceMode::Automatic,
            loop_sound: true,
            ..Default::default()
        };
        let flags = encode_transition_flags(&transition);
        assert_eq!(flags & 0x40, 0x40);
    }

    #[test]
    fn test_write_transition_minimal() {
        let transition = TransitionInfo::default();
        let data = write_transition(&transition);

        assert!(data.len() >= 8);
    }
}
