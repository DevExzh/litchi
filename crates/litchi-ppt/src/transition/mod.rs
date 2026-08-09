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
