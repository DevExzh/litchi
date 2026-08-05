//! Archive-free Keynote semantic models.
//!
//! This crate owns the presentation vocabulary used by Keynote readers and
//! editors.  Archive objects, protobuf messages, package identifiers, and
//! mutation transactions remain in the concrete format crate.

#![forbid(unsafe_code)]

mod error;
mod time;

pub mod background;
pub mod build;
pub mod document;
pub mod show;
pub mod slide;
pub mod transition;

pub use background::{Angle, Background, Gradient, Kind, Opaque, Stop};
pub use build::{AnimationType, Build};
pub use document::Document;
pub use error::{Error, Result};
pub use show::{Mode, Settings, Show, Size};
pub use slide::{Slide, Transition};
pub use time::Seconds;
pub use transition::Effect;
