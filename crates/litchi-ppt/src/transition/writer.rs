//! Slide transition writer.
//!
//! Writes `PowerPoint` binary slide transition records.

use super::types::{
    AdvanceMode, TransitionDirection, TransitionInfo, TransitionSpeed, TransitionType,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};

/// Write `SSSlideInfoAtom` record with transition information.
///
/// # Errors
///
/// Returns an invalid-format error when the requested type/direction pair has
/// no representation in the `[MS-PPT]` 2.6.6 transition table.
pub fn write_transition(transition: &TransitionInfo) -> Result<Vec<u8>> {
    let mut data = Vec::new();

    let Some([effect_direction, effect_type, effect_speed]) = super::encode_visual(
        transition.transition_type,
        transition.direction,
        transition.speed,
    ) else {
        return Err(Error::InvalidFormat(
            "transition type/direction is not representable by [MS-PPT] 2.6.6".into(),
        ));
    };

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
    data.push(effect_direction);

    // effectType is 1 byte, not 2!
    data.push(effect_type);

    // effectTransitionFlags (2 bytes)
    let flags = encode_transition_flags(transition);
    data.extend(&flags.to_le_bytes());

    // speed (1 byte)
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

    Ok(result)
}

/// Encode transition type to effect type value (1 byte).
///
/// Values follow the exact effect table in `[MS-PPT]` 2.6.6. Effects which
/// exist in newer presentation formats but have no binary-PPT representation
/// return `None` and are rejected by the writer.
pub(super) const fn encode_transition_type(transition_type: TransitionType) -> Option<u8> {
    match transition_type {
        TransitionType::None | TransitionType::Cut => Some(0),
        TransitionType::Random => Some(1),
        TransitionType::Blinds => Some(2),
        TransitionType::Checkerboard => Some(3),
        TransitionType::Cover => Some(4),
        TransitionType::Dissolve => Some(5),
        TransitionType::Fade => Some(6),
        TransitionType::Uncover => Some(7),
        TransitionType::RandomBars => Some(8),
        TransitionType::Strips => Some(9),
        TransitionType::Wipe => Some(10),
        TransitionType::Box => Some(11),
        TransitionType::Split => Some(13),
        TransitionType::Diamond => Some(17),
        TransitionType::Plus => Some(18),
        TransitionType::Wedge => Some(19),
        TransitionType::Push => Some(20),
        TransitionType::Comb => Some(21),
        TransitionType::Newsflash => Some(22),
        TransitionType::AlphaFade => Some(23),
        TransitionType::Wheel => Some(26),
        TransitionType::Circle => Some(27),
        TransitionType::Undefined => Some(255),
        TransitionType::Zoom
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
        | TransitionType::Morph => None,
    }
}

/// Encode transition direction based on type.
///
/// Direction values follow MS-PPT 2.6.6 (`SlideShowSlideInfoAtom`) and are the
/// exact inverse of `parse_transition_direction` in the reader.
pub(super) const fn encode_transition_direction(
    direction: TransitionDirection,
    transition_type: TransitionType,
) -> Option<u8> {
    // Each nested match is intentionally a compact whitelist for one wire
    // effect; all remaining enum directions are unrepresentable for it.
    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "nested matches whitelist the exact directions permitted by each MS-PPT effect"
    )]
    match transition_type {
        TransitionType::Cut => match direction {
            TransitionDirection::ThroughBlack => Some(1),
            _ => None,
        },
        TransitionType::Blinds => match direction {
            TransitionDirection::Vertical => Some(0),
            TransitionDirection::Horizontal => Some(1),
            _ => None,
        },
        TransitionType::Checkerboard | TransitionType::RandomBars | TransitionType::Comb => {
            match direction {
                TransitionDirection::Horizontal => Some(0),
                TransitionDirection::Vertical => Some(1),
                _ => None,
            }
        },
        TransitionType::Cover | TransitionType::Uncover => match direction {
            TransitionDirection::FromLeft => Some(0),
            TransitionDirection::FromTop => Some(1),
            TransitionDirection::FromRight => Some(2),
            TransitionDirection::FromBottom => Some(3),
            TransitionDirection::LeftUp => Some(4),
            TransitionDirection::RightUp => Some(5),
            TransitionDirection::LeftDown => Some(6),
            TransitionDirection::RightDown => Some(7),
            _ => None,
        },
        TransitionType::Strips => match direction {
            TransitionDirection::LeftUp => Some(4),
            TransitionDirection::RightUp => Some(5),
            TransitionDirection::LeftDown => Some(6),
            TransitionDirection::RightDown => Some(7),
            _ => None,
        },
        TransitionType::Wipe | TransitionType::Push => match direction {
            TransitionDirection::FromLeft => Some(0),
            TransitionDirection::FromTop => Some(1),
            TransitionDirection::FromRight => Some(2),
            TransitionDirection::FromBottom => Some(3),
            _ => None,
        },
        TransitionType::Box => match direction {
            TransitionDirection::Out => Some(0),
            TransitionDirection::In => Some(1),
            _ => None,
        },
        TransitionType::Split => match direction {
            TransitionDirection::HorizontalOut => Some(0),
            TransitionDirection::HorizontalIn => Some(1),
            TransitionDirection::VerticalOut => Some(2),
            TransitionDirection::VerticalIn => Some(3),
            _ => None,
        },
        TransitionType::Wheel => match direction {
            TransitionDirection::Spokes1 => Some(1),
            TransitionDirection::Spokes2 => Some(2),
            TransitionDirection::Spokes3 => Some(3),
            TransitionDirection::Spokes4 => Some(4),
            TransitionDirection::Spokes8 => Some(8),
            _ => None,
        },
        TransitionType::None
        | TransitionType::Random
        | TransitionType::Dissolve
        | TransitionType::Fade
        | TransitionType::Wedge
        | TransitionType::Diamond
        | TransitionType::Plus
        | TransitionType::Newsflash
        | TransitionType::AlphaFade
        | TransitionType::Circle
        | TransitionType::Undefined => match direction {
            TransitionDirection::None => Some(0),
            _ => None,
        },
        TransitionType::Zoom
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
        | TransitionType::Morph => None,
    }
}

/// Encode transition speed.
pub(super) fn encode_transition_speed(speed: TransitionSpeed) -> u8 {
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
        assert_eq!(encode_transition_type(TransitionType::None), Some(0));
        assert_eq!(encode_transition_type(TransitionType::Random), Some(1));
        assert_eq!(encode_transition_type(TransitionType::Cover), Some(4));
        assert_eq!(encode_transition_type(TransitionType::Box), Some(11));
        assert_eq!(encode_transition_type(TransitionType::Diamond), Some(17));
        assert_eq!(encode_transition_type(TransitionType::Morph), None);
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
        let data = write_transition(&transition).unwrap();

        assert!(data.len() >= 8);
    }
}
