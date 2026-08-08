//! Archive-free Keynote semantic models.
//!
//! This crate owns the presentation vocabulary used by Keynote readers and
//! editors.  Archive objects, protobuf messages, package identifiers, and
//! mutation transactions remain in the concrete format crate.
//!
//! # Edit slide playback state
//!
//! Select slides by an exact navigator name or a checked semantic position;
//! native IWA identifiers never enter the public transaction.
//!
//! ```no_run
//! use std::fs::OpenOptions;
//! use std::io::Write as _;
//!
//! use litchi_keynote::{Package, SlideSelector};
//!
//! let package = Package::open("input.key")?;
//! let mut edit = package.edit();
//! edit.skip_slide(SlideSelector::name("Appendix"))?;
//! let commit = edit.commit()?;
//! let mut output = OpenOptions::new()
//!     .write(true)
//!     .create_new(true)
//!     .open("output.key")?;
//! output.write_all(commit.package().source_bytes())?;
//! output.sync_all()?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

#![forbid(unsafe_code)]

mod error;
mod package;
mod selector;
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
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub use package::__semantic_document_from_prepared_source;
pub use package::{
    Commit, Diagnostics, Edit, EditError, Limits, MAX_OBJECTS, MAX_REFERENCES, MAX_SLIDES,
    MAX_TEXT_BYTES, MAX_TEXT_FRAGMENTS, MAX_TEXT_STORAGES, Package, Patch, PayloadLimitKind,
    ReadError, ReadOptions, SemanticLimitKind, SemanticLimits, SemanticLimitsError, SemanticPath,
    Stats, TextStorageFailure,
};
pub use selector::{SlideSelector, SlideSelectorError, SlideSelectorResult};
pub use show::{Mode, Settings, Show, Size};
pub use slide::media::MovieKind;
pub use slide::{Slide, Transition};
pub use time::Seconds;
pub use transition::Effect;
