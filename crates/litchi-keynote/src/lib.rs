//! Archive-free Keynote semantic models.
//!
//! This crate owns the presentation vocabulary used by Keynote readers and
//! editors.  Archive objects, protobuf messages, package identifiers, and
//! mutation transactions remain in the concrete format crate.

#![forbid(unsafe_code)]

mod error;
mod package;
mod time;

pub mod background;
pub mod build;
pub mod chart;
pub mod document;
pub mod show;
pub mod slide;
pub mod soundtrack;
pub mod transition;

pub use background::{Angle, Background, Gradient, Kind, Opaque, Stop};
pub use build::{AnimationType, Build};
pub use chart::ChartSelector;
pub use document::Document;
pub use error::{Error, Result};
pub use package::{Limits, Package, ReadError, Stats};
pub use show::{Mode, Settings, Show, Size};
pub use slide::media::MovieKind;
pub use slide::{Slide, Transition};
pub use time::Seconds;
pub use transition::Effect;
