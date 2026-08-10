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

#[allow(
    clippy::module_name_repetitions,
    reason = "`parse_transition` is the established public entry point of this module; renaming it would break downstream crates"
)]
pub use parser::parse_transition;
#[allow(
    clippy::module_name_repetitions,
    reason = "`TransitionSound` is the established public API name; renaming it would break downstream crates"
)]
pub use sound::TransitionSound;
#[allow(
    clippy::module_name_repetitions,
    reason = "the `Transition*` type names are the established public API of this module; renaming them would break downstream crates"
)]
pub use types::{
    AdvanceMode, SoundAction, TransitionDirection, TransitionInfo, TransitionSpeed, TransitionType,
};
#[allow(
    clippy::module_name_repetitions,
    reason = "`write_transition` is the established public entry point of this module; renaming it would break downstream crates"
)]
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
