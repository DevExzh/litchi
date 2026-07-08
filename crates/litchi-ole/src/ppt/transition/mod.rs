//! PPT slide transition support.
//!
//! This module provides structures and functions for parsing and writing
//! PowerPoint binary slide transition records, including:
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
