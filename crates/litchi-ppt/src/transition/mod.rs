//! PPT slide transition support.
//!
//! This module provides structures and functions for parsing and writing
//! `PowerPoint` binary slide transition records, including:
//! - Transition types and effects
//! - Transition speeds and directions
//! - Slide advance modes (click, automatic, both)
//! - Sound support for transitions

pub mod parser;
pub mod sound;
pub mod types;
pub mod writer;

pub use parser::parse_transition;
pub use sound::TransitionSound;
pub use types::{
    AdvanceMode, SoundAction, TransitionDirection, TransitionInfo, TransitionSpeed, TransitionType,
};
pub use writer::write_transition;

pub(crate) fn decode_visual(
    bytes: [u8; 3],
) -> Option<(TransitionType, TransitionDirection, TransitionSpeed)> {
    let (transition_type, direction) =
        parser::parse_transition_visual(u16::from(bytes[1]), bytes[0]);
    let speed = parser::parse_transition_speed(bytes[2]);
    let canonical = [
        writer::encode_transition_direction(direction, transition_type)?,
        writer::encode_transition_type(transition_type)?,
        writer::encode_transition_speed(speed),
    ];
    (canonical == bytes).then_some((transition_type, direction, speed))
}

pub(crate) fn encode_visual(
    transition_type: TransitionType,
    direction: TransitionDirection,
    speed: TransitionSpeed,
) -> Option<[u8; 3]> {
    let bytes = [
        writer::encode_transition_direction(direction, transition_type)?,
        writer::encode_transition_type(transition_type)?,
        writer::encode_transition_speed(speed),
    ];
    (decode_visual(bytes) == Some((transition_type, direction, speed))).then_some(bytes)
}
