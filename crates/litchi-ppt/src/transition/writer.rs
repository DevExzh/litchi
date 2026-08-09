//! Slide transition writer.
//!
//! Writes `PowerPoint` binary slide transition records.

use super::types::{
    AdvanceMode, TransitionDirection, TransitionInfo, TransitionSpeed, TransitionType,
};
use crate::consts::RecordType;

/// Write `SSSlideInfoAtom` record with transition information.
#[must_use]
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

    let sound_id_ref = u32::from(transition.sound.is_some());
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

    #[allow(
        clippy::cast_possible_truncation,
        reason = "the SSSlideInfoAtom payload built above is always 16 bytes, so its length fits in `u32`"
    )]
    let header = create_record_header(RecordType::SSSlideInfoAtom, 0x00, 0, data.len() as u32);

    let mut result = Vec::new();
    result.extend(header);
    result.extend(data);

    result
}

/// Encode transition type to effect type value (1 byte).
///
/// These are the crate's established `SlideShowSlideInfoAtom` effect values,
/// shared with `parse_transition_type` in the reader so authored files
/// round-trip through this crate unchanged.
fn encode_transition_type(transition_type: TransitionType) -> u8 {
    match transition_type {
        TransitionType::None => 0,
        TransitionType::Blinds => 1,
        TransitionType::Checkerboard => 2,
        TransitionType::Cover => 3,
        TransitionType::Dissolve => 4,
        TransitionType::Fade => 5,
        TransitionType::Uncover => 6,
        TransitionType::RandomBars => 7,
        TransitionType::Strips => 8,
        TransitionType::Wipe => 9,
        TransitionType::Box => 10,
        TransitionType::Random => 11,
        TransitionType::Zoom => 20,
        TransitionType::Split => 13,
        TransitionType::Cut => 17, // DIAMOND
        TransitionType::Push => 18,
        TransitionType::Comb => 19,
        TransitionType::Wedge => 21,
        TransitionType::Wheel => 22,
        TransitionType::Newsflash => 23,
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
///
/// Direction values follow MS-PPT 2.6.6 (`SlideShowSlideInfoAtom`) and are the
/// exact inverse of `parse_transition_direction` in the reader.
fn encode_transition_direction(
    direction: TransitionDirection,
    transition_type: TransitionType,
) -> u8 {
    match transition_type {
        TransitionType::Blinds => match direction {
            // MS-PPT 2.6.6 Blinds: 0=Vertical, 1=Horizontal
            TransitionDirection::Horizontal => 1,
            TransitionDirection::None
            | TransitionDirection::Vertical
            | TransitionDirection::FromLeft
            | TransitionDirection::FromRight
            | TransitionDirection::FromTop
            | TransitionDirection::FromBottom
            | TransitionDirection::In
            | TransitionDirection::Out
            | TransitionDirection::LeftDown
            | TransitionDirection::LeftUp
            | TransitionDirection::RightDown
            | TransitionDirection::RightUp => 0,
        },
        TransitionType::Checkerboard | TransitionType::RandomBars => match direction {
            // Checkerboard: 0=horizontal, 1=vertical
            TransitionDirection::Vertical => 1,
            TransitionDirection::None
            | TransitionDirection::Horizontal
            | TransitionDirection::FromLeft
            | TransitionDirection::FromRight
            | TransitionDirection::FromTop
            | TransitionDirection::FromBottom
            | TransitionDirection::In
            | TransitionDirection::Out
            | TransitionDirection::LeftDown
            | TransitionDirection::LeftUp
            | TransitionDirection::RightDown
            | TransitionDirection::RightUp => 0,
        },
        TransitionType::Split => match direction {
            // MS-PPT 2.6.6 Split: 0=Horizontally out, 2=Vertically out
            // (the in/out axis is not representable in `TransitionDirection`)
            TransitionDirection::Vertical => 2,
            // Horizontal split (opens vertically)
            TransitionDirection::None
            | TransitionDirection::Horizontal
            | TransitionDirection::FromLeft
            | TransitionDirection::FromRight
            | TransitionDirection::FromTop
            | TransitionDirection::FromBottom
            | TransitionDirection::In
            | TransitionDirection::Out
            | TransitionDirection::LeftDown
            | TransitionDirection::LeftUp
            | TransitionDirection::RightDown
            | TransitionDirection::RightUp => 0,
        },
        TransitionType::Cover
        | TransitionType::Uncover
        | TransitionType::Wipe
        | TransitionType::Push => match direction {
            // MS-PPT 2.6.6 Cover/Uncover/Wipe/Push: 0=Left, 1=Up, 2=Right, 3=Down
            TransitionDirection::FromTop => 1,
            TransitionDirection::FromRight => 2,
            TransitionDirection::FromBottom => 3,
            TransitionDirection::None
            | TransitionDirection::Horizontal
            | TransitionDirection::Vertical
            | TransitionDirection::FromLeft
            | TransitionDirection::In
            | TransitionDirection::Out
            | TransitionDirection::LeftDown
            | TransitionDirection::LeftUp
            | TransitionDirection::RightDown
            | TransitionDirection::RightUp => 0,
        },
        TransitionType::Strips => match direction {
            // MS-PPT 2.6.6 Strips: 4=Left Up, 5=Right Up, 6=Left Down, 7=Right Down
            TransitionDirection::LeftUp => 4,
            TransitionDirection::RightUp => 5,
            TransitionDirection::LeftDown => 6,
            TransitionDirection::None
            | TransitionDirection::Horizontal
            | TransitionDirection::Vertical
            | TransitionDirection::FromLeft
            | TransitionDirection::FromRight
            | TransitionDirection::FromTop
            | TransitionDirection::FromBottom
            | TransitionDirection::In
            | TransitionDirection::Out
            | TransitionDirection::RightDown => 7,
        },
        TransitionType::Box | TransitionType::Zoom => match direction {
            // MS-PPT 2.6.6 Box In/Out: 0=Out, 1=In
            TransitionDirection::Out => 0,
            TransitionDirection::None
            | TransitionDirection::Horizontal
            | TransitionDirection::Vertical
            | TransitionDirection::FromLeft
            | TransitionDirection::FromRight
            | TransitionDirection::FromTop
            | TransitionDirection::FromBottom
            | TransitionDirection::In
            | TransitionDirection::LeftDown
            | TransitionDirection::LeftUp
            | TransitionDirection::RightDown
            | TransitionDirection::RightUp => 1,
        },
        TransitionType::None
        | TransitionType::Cut
        | TransitionType::Dissolve
        | TransitionType::Fade
        | TransitionType::Comb
        | TransitionType::Wheel
        | TransitionType::Wedge
        | TransitionType::Random
        | TransitionType::Newsflash
        | TransitionType::Vortex
        | TransitionType::Shred
        | TransitionType::Switch
        | TransitionType::Flip
        | TransitionType::Gallery
        | TransitionType::Cube
        | TransitionType::Doors
        | TransitionType::Window
        | TransitionType::Ferris
        | TransitionType::Conveyor
        | TransitionType::Rotate
        | TransitionType::Pan
        | TransitionType::Glitter
        | TransitionType::Honeycomb
        | TransitionType::Flash
        | TransitionType::Ripple
        | TransitionType::Fracture
        | TransitionType::Crush
        | TransitionType::Peel
        | TransitionType::PageCurl
        | TransitionType::Airplane
        | TransitionType::Origami
        | TransitionType::Morph => 0,
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
///
/// Bit layout per MS-PPT 2.6.6 (`SlideShowSlideInfoAtom`):
/// - bit 0: `fManualAdvance`
/// - bit 2: `fHidden`
/// - bit 4: `fSound`
/// - bit 6: `fLoopSound`
/// - bit 8: `fStopSound`
/// - bit 10: `fAutoAdvance`
/// - bit 12: `fCursorVisible`
fn encode_transition_flags(transition: &TransitionInfo) -> u16 {
    let mut flags = 0u16;

    // Manual advance (on click)
    if matches!(
        transition.advance_mode,
        AdvanceMode::OnClick | AdvanceMode::Both
    ) {
        flags |= 1 << 0; // fManualAdvance
    }

    // Auto advance
    if matches!(
        transition.advance_mode,
        AdvanceMode::Automatic | AdvanceMode::Both
    ) {
        flags |= 1 << 10; // fAutoAdvance
    }

    if transition.sound.is_some() {
        flags |= 1 << 4; // fSound
    }

    if transition.loop_sound {
        flags |= 1 << 6; // fLoopSound
    }

    flags
}

/// Create a PPT record header.
fn create_record_header(
    record_type: RecordType,
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
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
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
